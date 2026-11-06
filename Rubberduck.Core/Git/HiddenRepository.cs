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

    public enum GitMergeKind
    {
        /// <summary>The remote had nothing this side lacks.</summary>
        UpToDate,
        FastForwarded,
        Merged,

        /// <summary>The two sides changed the same lines. The merge was undone.</summary>
        Conflicts
    }

    public class GitMergeOutcome
    {
        public GitMergeOutcome(GitMergeKind kind, IReadOnlyList<GitChange> changes, IReadOnlyList<string> conflicts)
        {
            Kind = kind;
            Changes = changes;
            Conflicts = conflicts;
        }

        public GitMergeKind Kind { get; }

        /// <summary>What the merge changed in the work tree.</summary>
        public IReadOnlyList<GitChange> Changes { get; }

        /// <summary>The files that conflicted, for <see cref="GitMergeKind.Conflicts"/>.</summary>
        public IReadOnlyList<string> Conflicts { get; }
    }

    public class GitCommit
    {
        public string Id { get; set; }
        public string Author { get; set; }
        public DateTimeOffset Date { get; set; }
        public string Subject { get; set; }

        /// <summary>For the history of one file: where the file was in this commit (it may have been moved or renamed since). Null otherwise.</summary>
        public string Path { get; set; }
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

        /// <summary>
        /// The commits that added or changed one file, newest first, following it across renames and moves; each says where the file was then.
        /// </summary>
        public IReadOnlyList<GitCommit> LogOf(string relativePath, int max = 100)
        {
            if (!_git.Run(Prefix().Concat(new[] { "rev-parse", "--verify", "--quiet", "HEAD" }), WorkTree).Ok)
            {
                return new GitCommit[0];
            }

            var output = Run("log", "--follow", "--diff-filter=AMR", "-n", max.ToString(), "--format=%x1e%H%x1f%an%x1f%aI%x1f%s", "--name-only", "--", PathOf(relativePath)).Output;
            var commits = new List<GitCommit>();
            foreach (var record in output.Split(new[] { '\x1e' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var lines = record.Split('\n').Select(line => line.Trim('\r')).ToList();
                var fields = lines[0].Split('\x1f');
                if (fields.Length != 4)
                {
                    continue;
                }

                commits.Add(new GitCommit
                {
                    Id = fields[0], Author = fields[1], Date = DateTimeOffset.Parse(fields[2]), Subject = fields[3],
                    Path = lines.Skip(1).FirstOrDefault(line => line.Length > 0)
                });
            }

            return commits;
        }

        /// <summary>A file as it was in a commit, or null if the commit has no such file.</summary>
        public byte[] TryReadFileAt(string commit, string relativePath)
        {
            var result = _git.Run(Prefix().Concat(new[] { "show", CommitOf(commit) + ":" + PathOf(relativePath) }), WorkTree);
            return result.Ok ? result.OutputBytes : null;
        }

        /// <summary>A file as it was in a commit, byte for byte.</summary>
        public byte[] ReadFileAt(string commit, string relativePath)
            => Run("show", CommitOf(commit) + ":" + PathOf(relativePath)).OutputBytes;

        /// <summary>Puts a file back as it was in a commit. Only the work tree changes; nothing is committed.</summary>
        public void RestoreFile(string commit, string relativePath)
            => Run("checkout", CommitOf(commit), "--", PathOf(relativePath));

        /// <summary>The same repository, run with another runner: the one that carries a remote's credentials for the call.</summary>
        public HiddenRepository WithRunner(GitRunner runner) => new HiddenRepository(runner, _fileSystem, Root);

        // ---- Remote ----

        public const string RemoteName = "origin";

        /// <summary>The URL of <c>origin</c>, or null if the repository has none.</summary>
        public string RemoteUrl()
        {
            var result = _git.Run(Prefix().Concat(new[] { "remote", "get-url", RemoteName }), WorkTree);
            return result.Ok ? result.Output.Trim() : null;
        }

        /// <summary>Points <c>origin</c> at the URL, adding it if there is none. The URL is not checked here; what may be used is the caller's decision.</summary>
        public void SetRemote(string url)
        {
            if (RemoteUrl() == null)
            {
                Run("remote", "add", RemoteName, url);
            }
            else
            {
                Run("remote", "set-url", RemoteName, url);
            }
        }

        public string CurrentBranch() => Run("symbolic-ref", "--short", "HEAD").Output.Trim();

        public void Fetch() => Run("fetch", "--quiet", RemoteName);

        /// <summary>Sends the current branch to <c>origin</c> and makes it the branch's upstream. Never forces: a remote that has moved on is refused.</summary>
        public void Push() => Run("push", "--quiet", "--set-upstream", RemoteName, CurrentBranch());

        /// <summary>
        /// How far the branch is from its copy on the remote as last fetched: commits only here, and commits only there.
        /// Null if the remote has no such branch (an empty remote, or nothing fetched yet).
        /// </summary>
        public (int ahead, int behind)? AheadBehind()
        {
            var remote = "refs/remotes/" + RemoteName + "/" + CurrentBranch();
            if (!_git.Run(Prefix().Concat(new[] { "rev-parse", "--verify", "--quiet", remote }), WorkTree).Ok)
            {
                return null;
            }

            var counts = Run("rev-list", "--left-right", "--count", "HEAD..." + remote).Output.Trim().Split('\t', ' ');
            return (int.Parse(counts[0]), int.Parse(counts[1]));
        }

        /// <summary>
        /// Brings what the remote has into the branch: nothing if it has nothing new, a fast-forward if this side has nothing of its own, otherwise a merge.
        /// A merge that would conflict is undone and reported, so the repository is left as it was.
        /// </summary>
        public GitMergeOutcome MergeRemote(string authorName, string authorEmail)
        {
            var counts = AheadBehind();
            if (counts == null || counts.Value.behind == 0)
            {
                return new GitMergeOutcome(GitMergeKind.UpToDate, new GitChange[0], new string[0]);
            }

            var before = Run("rev-parse", "HEAD").Output.Trim();
            var remote = RemoteName + "/" + CurrentBranch();
            var fastForward = counts.Value.ahead == 0;
            var merge = _git.Run(Prefix().Concat(new[]
            {
                "-c", "user.name=" + authorName, "-c", "user.email=" + authorEmail, "-c", "commit.gpgsign=false",
                "merge", "--no-edit", "--quiet", fastForward ? "--ff-only" : "--no-ff", remote
            }), WorkTree);

            if (!merge.Ok)
            {
                var conflicts = _git.Run(Prefix().Concat(new[] { "diff", "--name-only", "--diff-filter=U", "-z" }), WorkTree).Output
                    .Split(new[] { '\0' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                _git.Run(Prefix().Concat(new[] { "merge", "--abort" }), WorkTree);
                if (conflicts.Count == 0)
                {
                    throw new GitException("git merge failed: " + merge.Error, merge.ExitCode);
                }

                return new GitMergeOutcome(GitMergeKind.Conflicts, new GitChange[0], conflicts);
            }

            return new GitMergeOutcome(fastForward ? GitMergeKind.FastForwarded : GitMergeKind.Merged, ChangesSince(before), new string[0]);
        }

        //What a merge changed in the work tree, as paths and kinds.
        private IReadOnlyList<GitChange> ChangesSince(string commit)
        {
            var fields = Run("diff", "--name-status", "-z", "--no-renames", CommitOf(commit), "HEAD").Output
                .Split(new[] { '\0' }, StringSplitOptions.RemoveEmptyEntries);
            var changes = new List<GitChange>();
            for (var i = 0; i + 1 < fields.Length; i += 2)
            {
                var kind = fields[i] == "A" ? GitChangeKind.Added : fields[i] == "D" ? GitChangeKind.Deleted : GitChangeKind.Modified;
                changes.Add(new GitChange(fields[i + 1], kind));
            }

            return changes;
        }

        /// <summary>Everything that differs from the last commit as one patch (the commit would hold exactly this). The bookkeeping file is left out.</summary>
        public string StagedDiff()
        {
            Run("add", "--all");
            var excluded = LocalOnlyFiles.Select(file => ":(exclude)" + file);
            return Run(new[] { "diff", "--cached", "--no-color", "--no-ext-diff", "--no-renames", "--unified=3", "--", "." }.Concat(excluded).ToArray()).Output;
        }

        /// <summary>Everything that differs from the last commit as one patch that <c>git apply</c> can take, binary files included, read byte for byte (see <see cref="GitHunkParser.Bytes"/>).</summary>
        public string StagedPatch()
        {
            Run("add", "--all");
            var excluded = LocalOnlyFiles.Select(file => ":(exclude)" + file);
            var result = Run(new[] { "diff", "--cached", "--binary", "--no-color", "--no-ext-diff", "--no-renames", "--unified=3", "--", "." }.Concat(excluded).ToArray());
            return GitHunkParser.Bytes.GetString(result.OutputBytes);
        }

        /// <summary>
        /// Commits the changes as several commits, each holding the hunks of its part. All of them or none: if a part cannot be applied,
        /// or the parts do not add up to every change, the history is put back as it was and nothing of the work tree is touched.
        /// </summary>
        /// <returns>The new commits' ids, in order.</returns>
        public IReadOnlyList<string> CommitParts(IReadOnlyList<GitCommitPart> parts, string authorName, string authorEmail)
        {
            if (parts.Any(part => string.IsNullOrWhiteSpace(part.Message) || part.Hunks.Count == 0))
            {
                throw new ArgumentException("Every commit needs a message and at least one change.", nameof(parts));
            }

            var before = _git.Run(Prefix().Concat(new[] { "rev-parse", "--verify", "--quiet", "HEAD" }), WorkTree);
            var original = before.Ok ? before.Output.Trim() : null;
            var patchFile = _fileSystem.Path.Combine(Root, "part.patch");
            var ids = new List<string>();
            try
            {
                foreach (var part in parts)
                {
                    //The index is what the next commit holds: start each part from the last commit.
                    if (original != null || ids.Count > 0)
                    {
                        Run("reset", "--quiet");
                    }
                    else
                    {
                        Run("read-tree", "--empty");
                    }

                    _fileSystem.File.WriteAllBytes(patchFile, GitHunkParser.Bytes.GetBytes(GitHunkParser.BuildPatch(part.Hunks)));
                    Run("apply", "--cached", "--recount", "--whitespace=nowarn", "--", patchFile);
                    _git.RunOrThrow(Prefix().Concat(new[]
                    {
                        "-c", "user.name=" + authorName, "-c", "user.email=" + authorEmail, "-c", "commit.gpgsign=false",
                        "commit", "--quiet", "--no-verify", "--message", part.Message
                    }), WorkTree);
                    ids.Add(Run("rev-parse", "HEAD").Output.Trim());
                }

                Run("add", "--all");
                if (Status().Count > 0)
                {
                    throw new GitException("The commits do not add up to every change: " + string.Join(", ", Status().Select(change => change.Path)) + " would be left out.", -1);
                }

                return ids;
            }
            catch
            {
                if (original != null)
                {
                    _git.Run(Prefix().Concat(new[] { "reset", "--quiet", original }), WorkTree);
                }
                else if (ids.Count > 0)
                {
                    _git.Run(Prefix().Concat(new[] { "update-ref", "-d", "HEAD" }), WorkTree);
                    _git.Run(Prefix().Concat(new[] { "read-tree", "--empty" }), WorkTree);
                }

                throw;
            }
            finally
            {
                if (_fileSystem.File.Exists(patchFile))
                {
                    _fileSystem.File.Delete(patchFile);
                }
            }
        }

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
