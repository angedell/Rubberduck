using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.UI.CodeExplorer.GitSync;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Git
{
    /// <summary>The remote needs a token that is not saved, or the saved one was refused.</summary>
    public class GitAuthenticationRequiredException : GitClientException
    {
        public GitAuthenticationRequiredException(string host, string message) : base(message)
        {
            Host = host;
        }

        public string Host { get; }
    }

    public class GitRemoteStatus
    {
        public string Url { get; set; }

        /// <summary>The commits this side has and the remote has not, as of the last fetch; null if the remote has no such branch yet.</summary>
        public int? Ahead { get; set; }

        public int? Behind { get; set; }
    }

    public class GitPullOutcome
    {
        public GitMergeKind Kind { get; set; }
        public List<string> Conflicts { get; } = new List<string>();
        public List<string> Notes { get; } = new List<string>();
        public List<string> Errors { get; } = new List<string>();
        public string BackupFolder { get; set; }
    }

    public partial class GitClientService
    {
        /// <summary>Where the project's history is pushed to, or null.</summary>
        public string RemoteUrl(IVBProject project)
        {
            var repository = Open(project, null, out _);
            return repository.Exists ? repository.RemoteUrl() : null;
        }

        /// <summary>Points the project's history at a remote. Only https addresses: the token is the sign-in, and nothing else is run on the user's behalf.</summary>
        public void SetRemote(IVBProject project, string url)
        {
            url = (url ?? string.Empty).Trim();
            if (GitRemoteAuth.HostOf(url) == null)
            {
                throw new GitClientException("The address has to be an https one, like https://github.com/owner/repository.git, and carry no user name or password.");
            }

            var repository = Open(project, null, out _);
            repository.EnsureCreated();
            repository.SetRemote(url);
            Logger.Info($"Git client: remote of '{project.Name}' set to {url}.");
        }

        public bool HasToken(string host) => _credentials.Read(CredentialTargets.GitHost(host)) != null;

        public void SaveToken(string host, string user, string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                throw new GitClientException("The token is empty.");
            }

            _credentials.Write(CredentialTargets.GitHost(host), string.IsNullOrWhiteSpace(user) ? null : user.Trim(), token.Trim());
            Logger.Info($"Git client: a token for {host} was saved in the Credential Manager.");
        }

        public bool ForgetToken(string host) => _credentials.Delete(CredentialTargets.GitHost(host));

        /// <summary>The address of the remote, whatever the project; the token page the user needs depends on it.</summary>
        public string RemoteHost(IVBProject project) => GitRemoteAuth.HostOf(RemoteUrl(project));

        /// <summary>Fetches and says how far apart the two sides are.</summary>
        public GitRemoteStatus RemoteStatus(IVBProject project)
        {
            var (repository, _) = OpenRemote(project);
            repository.Fetch();
            var distance = repository.AheadBehind();
            return new GitRemoteStatus { Url = repository.RemoteUrl(), Ahead = distance?.ahead, Behind = distance?.behind };
        }

        /// <summary>Sends the commits to the remote. Never forced: if the remote has moved on, the user has to pull first.</summary>
        public GitRemoteStatus Push(IVBProject project)
        {
            var timer = Stopwatch.StartNew();
            var (repository, host) = OpenRemote(project);
            if (!repository.Exists || repository.Log(1).Count == 0)
            {
                throw new GitClientException("There is nothing to push yet: commit first.");
            }

            Guarded(host, repository.Fetch);
            var distance = repository.AheadBehind();
            if (distance != null && distance.Value.behind > 0)
            {
                throw new GitClientException($"The remote has {distance.Value.behind} commit(s) this side does not have. Pull first, then push.");
            }

            if (distance != null && distance.Value.ahead == 0)
            {
                throw new GitClientException("The remote already has every commit.");
            }

            Guarded(host, repository.Push);
            Logger.Info($"Git client: pushed '{project.Name}' to {repository.RemoteUrl()}; {timer.ElapsedMilliseconds}ms.");
            return new GitRemoteStatus { Url = repository.RemoteUrl(), Ahead = 0, Behind = 0 };
        }

        /// <summary>
        /// Brings the remote's commits into the project. The project is saved to the work tree first; if it holds changes that were not committed,
        /// nothing is pulled, because the files that come in could not be told apart from them. After a merge, the files that changed are imported in one batch.
        /// </summary>
        public GitPullOutcome Pull(IVBProject project, ISyncProgress progress = null)
        {
            var timer = Stopwatch.StartNew();
            var (repository, host) = OpenRemote(project);

            var snapshot = _snapshot.Take(project, repository.WorkTree, progress);
            if (snapshot.Errors.Count > 0)
            {
                throw new GitClientException("Nothing was pulled, because these modules could not be compared: " + string.Join("; ", snapshot.Errors));
            }

            var pending = repository.Status();
            if (pending.Count > 0)
            {
                throw new GitClientException($"The project has {pending.Count} change(s) that are not committed. Commit them first, so that nothing is lost.");
            }

            Guarded(host, repository.Fetch);
            var (name, email) = Identity(RunnerFor(host, out _));
            var merge = repository.MergeRemote(name, email);

            var outcome = new GitPullOutcome { Kind = merge.Kind };
            if (merge.Kind == GitMergeKind.Conflicts)
            {
                outcome.Conflicts.AddRange(merge.Conflicts);
                Logger.Info($"Git client: pull of '{project.Name}' stopped by {merge.Conflicts.Count} conflicting file(s): {string.Join(", ", merge.Conflicts)}.");
                return outcome;
            }

            var deleted = merge.Changes.Where(change => change.Kind == GitChangeKind.Deleted).Select(change => change.Path).ToList();
            foreach (var path in deleted)
            {
                outcome.Notes.Add($"'{path}' was deleted on the remote; the module is still in the project.");
            }

            var toImport = merge.Changes.Where(change => change.Kind != GitChangeKind.Deleted).Select(change => change.Path).ToList();
            if (toImport.Count > 0)
            {
                var result = _snapshot.ImportFiles(project, repository.WorkTree, toImport, progress);
                outcome.Notes.AddRange(result.Notes);
                outcome.Errors.AddRange(result.Errors);
                outcome.BackupFolder = result.BackupFolder;
            }

            Logger.Info($"Git client: pull of '{project.Name}' {merge.Kind}, {toImport.Count} file(s) imported, {deleted.Count} deleted on the remote, {outcome.Errors.Count} error(s); {timer.ElapsedMilliseconds}ms.");
            return outcome;
        }

        //The repository with a runner that signs in to its remote's host, if a token for it is saved.
        private (HiddenRepository repository, string host) OpenRemote(IVBProject project)
        {
            var repository = Open(project, null, out var plain);
            var url = repository.Exists ? repository.RemoteUrl() : null;
            if (url == null)
            {
                throw new GitClientException("This project has no remote yet. Set the remote first.");
            }

            var host = GitRemoteAuth.HostOf(url)
                ?? throw new GitClientException("Only https remotes are supported.");
            var signedIn = RunnerFor(host, out _, plain);
            return (repository.WithRunner(signedIn), host);
        }

        private GitRunner RunnerFor(string host, out bool hasToken, GitRunner basis = null)
        {
            var credential = _credentials.Read(CredentialTargets.GitHost(host));
            hasToken = credential != null;
            var auth = credential == null ? GitRemoteAuth.None : GitRemoteAuth.ForHost(host, credential);
            var executable = (basis ?? new GitRunner(_git.Find(_settings().GitPath).Path)).Executable;
            return new GitRunner(executable, TimeSpan.FromMinutes(5), auth.Environment, auth.Secrets);
        }

        //A remote that refuses the sign-in says so in words git prints; those become a request for a token rather than a stack of git output.
        private void Guarded(string host, Action action)
        {
            try
            {
                action();
            }
            catch (GitException exception) when (LooksLikeAuthenticationFailure(exception.Message))
            {
                var saved = HasToken(host);
                throw new GitAuthenticationRequiredException(host, saved
                    ? $"{host} refused the saved token. It may have expired or lack permission to write to the repository."
                    : $"{host} wants a sign-in. Enter a personal access token for it.");
            }
        }

        private static bool LooksLikeAuthenticationFailure(string message)
        {
            var text = message.ToLowerInvariant();
            return text.Contains("authentication failed") || text.Contains("could not read username") || text.Contains("terminal prompts disabled")
                   || text.Contains("returned error: 401") || text.Contains("returned error: 403") || text.Contains("invalid credentials");
        }
    }
}
