using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.Git
{
    public enum GitChangeKind
    {
        Added,
        Modified,
        Deleted,
        Renamed
    }

    public class GitChange
    {
        public GitChange(string path, GitChangeKind kind)
        {
            Path = path;
            Kind = kind;
        }

        /// <summary>Relative to the work tree, with forward slashes.</summary>
        public string Path { get; }

        public GitChangeKind Kind { get; }
    }

    public class GitCommit
    {
        public string Id { get; set; }
        public string Author { get; set; }
        public DateTimeOffset Date { get; set; }
        public string Subject { get; set; }
    }

    /// <summary>
    /// A repository that Rubberduck keeps out of the user's sight: the git directory and the work tree are two folders under <c>root</c>,
    /// and every call names both, so no <c>.git</c> is written to the work tree. Hooks never run, and line endings are never changed:
    /// VBA modules are CRLF, and a repository that "fixes" that would make every module look modified.
    /// </summary>
    public class HiddenRepository
    {
        private static readonly Regex CommitId = new Regex("^[0-9a-f]{7,64}$", RegexOptions.CultureInvariant);

        /// <summary>
        /// Files the work tree needs but the history does not: Git sync's bookkeeping, which changes with every export and means nothing in an old commit.
        /// They are kept out of the repository (the git directory's own <c>info/exclude</c>, so nothing is written to the work tree).
        /// </summary>
        public static readonly string[] LocalOnlyFiles = { "rubberduck.sync.json" };

        private readonly GitRunner _git;
        private readonly IFileSystem _fileSystem;

        public HiddenRepository(GitRunner git, IFileSystem fileSystem, string root)
        {
            _git = git;
            _fileSystem = fileSystem;
            Root = root;
        }

        public string Root { get; }
        public string GitDirectory => _fileSystem.Path.Combine(Root, "git");
        public string WorkTree => _fileSystem.Path.Combine(Root, "tree");

        //An empty folder for hooks, so that none of the repository's own ever runs.
        private string NoHooks => _fileSystem.Path.Combine(Root, "no-hooks");

        public bool Exists => _fileSystem.File.Exists(_fileSystem.Path.Combine(GitDirectory, "HEAD"));

        /// <summary>Creates the repository if it is not there yet. The work tree may be filled before or after.</summary>
        public void EnsureCreated()
        {
            _fileSystem.Directory.CreateDirectory(WorkTree);
            _fileSystem.Directory.CreateDirectory(NoHooks);
            if (!Exists)
            {
                _git.RunOrThrow(Prefix().Concat(new[] { "init", "--initial-branch=main", "--quiet" }), WorkTree);
            }

            KeepLocalOnlyFilesOut();
        }

        //Also repairs a repository that already holds them: they are dropped from the index (the file stays where it is) and leave the history with the next commit.
        private void KeepLocalOnlyFilesOut()
        {
            var exclude = _fileSystem.Path.Combine(GitDirectory, "info", "exclude");
            var current = _fileSystem.File.Exists(exclude) ? _fileSystem.File.ReadAllText(exclude) : string.Empty;
            var lines = current.Split(new[] { '\n' }).Select(line => line.Trim()).ToList();
            var missing = LocalOnlyFiles.Where(file => !lines.Contains("/" + file)).ToList();
            if (missing.Count > 0)
            {
                _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(exclude));
                var separator = current.Length == 0 || current.EndsWith("\n") ? string.Empty : "\n";
                _fileSystem.File.WriteAllText(exclude, current + separator + string.Concat(missing.Select(file => "/" + file + "\n")));
            }

            foreach (var file in LocalOnlyFiles)
            {
                if (_git.Run(Prefix().Concat(new[] { "ls-files", "--error-unmatch", "--", file }), WorkTree).Ok)
                {
                    Run("rm", "--cached", "--quiet", "--", file);
                }
            }
        }

        /// <summary>Everything that differs from the last commit, untracked files included.</summary>
        public IReadOnlyList<GitChange> Status()
        {
            var output = Run("status", "--porcelain=v1", "-z", "--untracked-files=all").Output;
            var records = output.Split(new[] { '\0' }, StringSplitOptions.RemoveEmptyEntries);
            var changes = new List<GitChange>();
            for (var i = 0; i < records.Length; i++)
            {
                var record = records[i];
                if (record.Length < 4)
                {
                    continue;
                }

                var index = record[0];
                var tree = record[1];
                var path = record.Substring(3);

                //A rename names the old path in the record after it.
                if (index == 'R' || index == 'C' || tree == 'R' || tree == 'C')
                {
                    i++;
                    changes.Add(new GitChange(path, GitChangeKind.Renamed));
                    continue;
                }

                //Still in the index from before it was kept out: it goes with the next commit, but is not news.
                if (!LocalOnlyFiles.Contains(path))
                {
                    changes.Add(new GitChange(path, KindOf(index, tree)));
                }
            }

            return changes;
        }

        private static GitChangeKind KindOf(char index, char tree)
        {
            if (index == '?' || index == 'A' || tree == 'A')
            {
                return GitChangeKind.Added;
            }

            return index == 'D' || tree == 'D' ? GitChangeKind.Deleted : GitChangeKind.Modified;
        }

        /// <summary>Commits everything that changed. Returns the new commit's id, or null if there was nothing to commit.</summary>
        public string Commit(string message, string authorName, string authorEmail)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                throw new ArgumentException("A commit needs a message.", nameof(message));
            }

            Run("add", "--all");
            if (Status().Count == 0)
            {
                return null;
            }

            _git.RunOrThrow(Prefix().Concat(new[]
            {
                "-c", "user.name=" + authorName, "-c", "user.email=" + authorEmail, "-c", "commit.gpgsign=false",
                "commit", "--quiet", "--no-verify", "--message", message
            }), WorkTree);
            return Run("rev-parse", "HEAD").Output.Trim();
        }

        /// <summary>The newest commits first. Empty for a repository without any.</summary>
        public IReadOnlyList<GitCommit> Log(int max = 100)
        {
            if (!_git.Run(Prefix().Concat(new[] { "rev-parse", "--verify", "--quiet", "HEAD" }), WorkTree).Ok)
            {
                return new GitCommit[0];
            }

            var output = Run("log", "-n", max.ToString(), "--format=%H%x1f%an%x1f%aI%x1f%s%x1e").Output;
            return output.Split('\x1e')
                .Select(record => record.Trim('\r', '\n').Split('\x1f'))
                .Where(fields => fields.Length == 4)
                .Select(fields => new GitCommit { Id = fields[0], Author = fields[1], Date = DateTimeOffset.Parse(fields[2]), Subject = fields[3] })
                .ToList();
        }

        /// <summary>A file as it was in a commit, byte for byte.</summary>
        public byte[] ReadFileAt(string commit, string relativePath)
            => Run("show", CommitOf(commit) + ":" + PathOf(relativePath)).OutputBytes;

        /// <summary>Puts a file back as it was in a commit. Only the work tree changes; nothing is committed.</summary>
        public void RestoreFile(string commit, string relativePath)
            => Run("checkout", CommitOf(commit), "--", PathOf(relativePath));

        private GitResult Run(params string[] args) => _git.RunOrThrow(Prefix().Concat(args), WorkTree);

        private IEnumerable<string> Prefix() => new[]
        {
            "--git-dir=" + GitDirectory,
            "--work-tree=" + WorkTree,
            "-c", "core.hooksPath=" + NoHooks,
            "-c", "core.autocrlf=false",
            "-c", "core.safecrlf=false",
            "-c", "core.quotepath=false",
            "-c", "core.fsmonitor=false",
            "-c", "protocol.ext.allow=never"
        };

        //Both end up on a command line next to options, so neither may look like one, or point out of the work tree.
        private static string CommitOf(string commit)
        {
            if (commit == null || !CommitId.IsMatch(commit))
            {
                throw new ArgumentException("Not a commit id: " + commit, nameof(commit));
            }

            return commit;
        }

        private static string PathOf(string relativePath)
        {
            var path = (relativePath ?? string.Empty).Replace('\\', '/');
            if (path.Length == 0 || path.StartsWith("-") || path.StartsWith("/") || path.Contains(":") || path.Split('/').Any(part => part == ".." || part == ".git"))
            {
                throw new ArgumentException("Not a path inside the work tree: " + relativePath, nameof(relativePath));
            }

            return path;
        }
    }
}
