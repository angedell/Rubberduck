using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using Rubberduck.Api;
using Rubberduck.Settings;
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
    public class GitClientService
    {
        private readonly IGitProvisioner _git;
        private readonly IFileSystem _fileSystem;
        private readonly IProjectSnapshot _snapshot;
        private readonly ProjectRepositories _repositories;
        private readonly Func<GeneralSettings> _settings;

        public GitClientService(IGitProvisioner git, IFileSystem fileSystem, IProjectSnapshot snapshot, ProjectRepositories repositories, Func<GeneralSettings> settings)
        {
            _git = git;
            _fileSystem = fileSystem;
            _snapshot = snapshot;
            _repositories = repositories;
            _settings = settings;
        }

        /// <summary>Saves the project as a commit.</summary>
        /// <param name="continueCopiedHistory">Settles a copied workbook: true continues the original's history, false starts an empty one, null asks (by throwing <see cref="ProjectCopiedException"/>).</param>
        /// <exception cref="GitClientException">The client is off, git is missing, the workbook was never saved, or some modules could not be written.</exception>
        public GitCommitOutcome Commit(IVBProject project, string message, bool? continueCopiedHistory = null)
        {
            var repository = Open(project, continueCopiedHistory, out var runner);
            repository.EnsureCreated();

            var result = _snapshot.Take(project, repository.WorkTree);
            if (result.Errors.Count > 0)
            {
                //A commit with a stale module in it would be a lie in the history.
                throw new GitClientException("Nothing was committed, because these modules could not be written: " + string.Join("; ", result.Errors));
            }

            var (name, email) = Identity(runner);
            var outcome = new GitCommitOutcome { CommitId = repository.Commit(message, name, email) };
            outcome.Notes.AddRange(result.Notes);
            return outcome;
        }

        /// <summary>What the project has changed since the last commit.</summary>
        public IReadOnlyList<GitChange> Changes(IVBProject project, bool? continueCopiedHistory = null)
        {
            var repository = Open(project, continueCopiedHistory, out _);
            repository.EnsureCreated();

            var result = _snapshot.Take(project, repository.WorkTree);
            if (result.Errors.Count > 0)
            {
                throw new GitClientException("These modules could not be compared: " + string.Join("; ", result.Errors));
            }

            return repository.Status();
        }

        /// <summary>The commits made for the project, newest first. Reads no module, so it is safe to ask at any time.</summary>
        public IReadOnlyList<GitCommit> History(IVBProject project, int max = 100)
        {
            var repository = Open(project, null, out _);
            return repository.Exists ? repository.Log(max) : new GitCommit[0];
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
