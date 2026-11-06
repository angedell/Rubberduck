using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Diagnostics;
using System.Linq;
using NLog;
using Rubberduck.Api;
using Rubberduck.Settings;
using Rubberduck.UI.CodeExplorer.GitSync;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Git
{
    /// <summary>Why the git client cannot do what was asked; the message is for the user.</summary>
    public class GitClientException : Exception
    {
        public GitClientException(string message) : base(message)
        {
        }
    }

    /// <summary>The workbook looks like a copy of another one that has a repository. The user decides whether the copy continues that history or starts its own.</summary>
    public class ProjectCopiedException : GitClientException
    {
        public ProjectCopiedException(string copyOfId, string copyOfPath)
            : base($"This workbook has the same name as '{copyOfPath}', which has a history. Is it a copy of it?")
        {
            CopyOfId = copyOfId;
            CopyOfPath = copyOfPath;
        }

        public string CopyOfId { get; }
        public string CopyOfPath { get; }
    }

    public class GitCommitOutcome
    {
        /// <summary>The new commit, or null if nothing had changed since the last one.</summary>
        public string CommitId { get; set; }

        public List<string> Notes { get; } = new List<string>();
    }

    /// <summary>
    /// The built-in git client for a project: finds its hidden repository, brings the repository's work tree up to date with the project through Git sync,
    /// and runs git on it. Everything that touches the VBE is Git sync's; git never sees it.
    /// </summary>
    public partial class GitClientService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IGitProvisioner _git;
        private readonly IFileSystem _fileSystem;
        private readonly IProjectSnapshot _snapshot;
        private readonly ProjectRepositories _repositories;
        private readonly Func<GeneralSettings> _settings;
        private readonly ICredentialStore _credentials;
        private readonly Func<string, string, IAiClient> _aiClient;

        public GitClientService(IGitProvisioner git, IFileSystem fileSystem, IProjectSnapshot snapshot, ProjectRepositories repositories, Func<GeneralSettings> settings,
            ICredentialStore credentials = null, Func<string, string, IAiClient> aiClient = null)
        {
            _aiClient = aiClient;
            _credentials = credentials ?? new WindowsCredentialStore();
            _git = git;
            _fileSystem = fileSystem;
            _snapshot = snapshot;
            _repositories = repositories;
            _settings = settings;
        }

        /// <summary>Whether the user switched the git client on in the settings.</summary>
        public bool Enabled => _settings().GitClientEnabled;

        /// <summary>Saves the project as a commit.</summary>
        /// <param name="continueCopiedHistory">Settles a copied workbook: true continues the original's history, false starts an empty one, null asks (by throwing <see cref="ProjectCopiedException"/>).</param>
        /// <exception cref="GitClientException">The client is off, git is missing, the workbook was never saved, or some modules could not be written.</exception>
        /// <param name="refresh">False when <see cref="Changes"/> has just brought the work tree up to date and nothing has changed since: the project is not read again.</param>
        public GitCommitOutcome Commit(IVBProject project, string message, bool? continueCopiedHistory = null, ISyncProgress progress = null, bool refresh = true)
        {
            var timer = Stopwatch.StartNew();
            var repository = Open(project, continueCopiedHistory, out var runner);
            repository.EnsureCreated();

            var result = refresh ? _snapshot.Take(project, repository.WorkTree, progress) : new GitSyncResult();
            var snapshotted = timer.ElapsedMilliseconds;
            if (result.Errors.Count > 0)
            {
                Logger.Warn($"Git client: not committing '{project.Name}', {result.Errors.Count} module(s) could not be written: {string.Join("; ", result.Errors)}");
                //A commit with a stale module in it would be a lie in the history.
                throw new GitClientException("Nothing was committed, because these modules could not be written: " + string.Join("; ", result.Errors));
            }

            var (name, email) = Identity(runner);
            var outcome = new GitCommitOutcome { CommitId = repository.Commit(message, name, email) };
            outcome.Notes.AddRange(result.Notes);
            Logger.Info($"Git client: commit of '{project.Name}' in {repository.Root}: {(outcome.CommitId ?? "nothing to commit")}; snapshot {snapshotted}ms (refresh {refresh}), total {timer.ElapsedMilliseconds}ms.");
            return outcome;
        }

        /// <summary>What the project has changed since the last commit.</summary>
        public IReadOnlyList<GitChange> Changes(IVBProject project, bool? continueCopiedHistory = null, ISyncProgress progress = null)
        {
            var timer = Stopwatch.StartNew();
            var repository = Open(project, continueCopiedHistory, out _);
            repository.EnsureCreated();

            var result = _snapshot.Take(project, repository.WorkTree, progress);
            if (result.Errors.Count > 0)
            {
                throw new GitClientException("These modules could not be compared: " + string.Join("; ", result.Errors));
            }

            var changes = repository.Status();
            Logger.Info($"Git client: '{project.Name}' has {changes.Count} change(s) since the last commit; {timer.ElapsedMilliseconds}ms.");
            return changes;
        }

        /// <summary>The commits made for the project, newest first. Reads no module, so it is safe to ask at any time.</summary>
        public IReadOnlyList<GitCommit> History(IVBProject project, int max = 100)
        {
            var repository = Open(project, null, out _);
            return repository.Exists ? repository.Log(max) : new GitCommit[0];
        }

        /// <summary>The commits in which the module changed, newest first. Reads no module; the module has to have been committed.</summary>
        public IReadOnlyList<GitCommit> ModuleHistory(IVBProject project, string moduleName, int max = 100)
        {
            var repository = Open(project, null, out _);
            return repository.Exists ? repository.LogOf(TreePathOf(repository, moduleName), max) : new GitCommit[0];
        }

        /// <summary>
        /// Replaces the module in the project with its version in a commit. The version it replaces is backed up first (<see cref="GitSyncResult.BackupFolder"/>),
        /// and nothing is committed: the restored module shows as a change, to be committed or thrown away like any other.
        /// </summary>
        public GitSyncResult RestoreModule(IVBProject project, string moduleName, GitCommit commit, ISyncProgress progress = null)
        {
            var timer = Stopwatch.StartNew();
            var repository = Open(project, null, out _);
            var treePath = TreePathOf(repository, moduleName);
            var pathThen = commit.Path ?? treePath;
            var content = repository.ReadFileAt(commit.Id, pathThen);

            //The module's file in the tree gets the old content, wherever the module is filed now; a form's binary goes with it.
            var target = _fileSystem.Path.Combine(repository.WorkTree, treePath.Replace('/', '\\'));
            _fileSystem.File.WriteAllBytes(target, content);
            var binary = repository.TryReadFileAt(commit.Id, _fileSystem.Path.ChangeExtension(pathThen, ".frx"));
            if (binary != null)
            {
                _fileSystem.File.WriteAllBytes(_fileSystem.Path.ChangeExtension(target, ".frx"), binary);
            }

            var result = _snapshot.Import(project, repository.WorkTree, moduleName, progress);
            Logger.Info($"Git client: restored '{moduleName}' of '{project.Name}' from {commit.Id.Substring(0, 7)} ({pathThen}); {result.Errors.Count} error(s); {timer.ElapsedMilliseconds}ms.");
            return result;
        }

        //Where Git sync last put the module in the tree: its own bookkeeping knows.
        private string TreePathOf(HiddenRepository repository, string moduleName)
        {
            var state = new GitSyncStateStore(_fileSystem).Load(_fileSystem.Path.Combine(repository.WorkTree, GitSyncState.FileName));
            if (!state.Modules.TryGetValue(moduleName, out var baseline) || string.IsNullOrEmpty(baseline.Path))
            {
                throw new GitClientException($"'{moduleName}' has not been committed yet, so it has no history to restore from.");
            }

            return baseline.Path;
        }

        private HiddenRepository Open(IVBProject project, bool? continueCopiedHistory, out GitRunner runner)
        {
            var settings = _settings();
            if (!settings.GitClientEnabled)
            {
                throw new GitClientException("The git client is switched off in the settings.");
            }

            var installation = _git.Find(settings.GitPath)
                ?? throw new GitClientException("git was not found. Switch the git client off and on again in the settings to download it, or enter the path of a git.exe.");
            runner = new GitRunner(installation.Path);

            var path = ProjectIds.PathOf(project);
            if (path == null)
            {
                throw new GitClientException("Save the workbook first: its history is kept by the file it is saved as.");
            }

            var resolution = _repositories.Resolve(path, project.Name);
            if (resolution.Outcome == ProjectRepositoryOutcome.CopyOf)
            {
                if (continueCopiedHistory == null)
                {
                    throw new ProjectCopiedException(resolution.CopyOfId, resolution.CopyOfPath);
                }

                resolution.Id = _repositories.Adopt(path, project.Name, resolution.CopyOfId, continueCopiedHistory.Value);
            }

            return new HiddenRepository(runner, _fileSystem, _repositories.RootOf(resolution.Id));
        }

        //The user's own git identity if they have one, otherwise who Windows says they are.
        private static (string name, string email) Identity(GitRunner runner)
        {
            string Configured(string key)
            {
                var result = runner.Run(new[] { "config", "--get", key });
                return result.Ok ? result.Output.Trim() : string.Empty;
            }

            var name = Configured("user.name");
            var email = Configured("user.email");
            return (name.Length > 0 ? name : Environment.UserName,
                email.Length > 0 ? email : Environment.UserName + "@" + Environment.MachineName + ".local");
        }
    }
}
