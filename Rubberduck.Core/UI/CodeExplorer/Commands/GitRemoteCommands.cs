using System;
using System.Collections.Generic;
using Rubberduck.Git;
using Rubberduck.Interaction;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.UI.CodeExplorer.GitClient;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    /// <summary>What the commands that talk to the remote have in common: when the host wants a token, ask for it once and go on.</summary>
    public abstract class GitRemoteCommandBase : GitClientCommandBase
    {
        protected GitRemoteCommandBase(IVbeEvents vbeEvents, IProjectsProvider projectsProvider, GitClientService git, IMessageBox messageBox)
            : base(vbeEvents, projectsProvider, git, messageBox)
        {
        }

        protected T SignedIn<T>(Func<T> work)
        {
            try
            {
                return work();
            }
            catch (GitAuthenticationRequiredException required)
            {
                var credential = GitTokenDialog.Ask(required.Host, required.Message);
                if (credential == null)
                {
                    throw new OperationCanceledException();
                }

                Git.SaveToken(required.Host, credential.User, credential.Secret);
                return work();
            }
        }
    }

    /// <summary>Says where the project's history is sent to.</summary>
    public class GitSetRemoteCommand : GitRemoteCommandBase
    {
        public GitSetRemoteCommand(IVbeEvents vbeEvents, IProjectsProvider projectsProvider, GitClientService git, IMessageBox messageBox)
            : base(vbeEvents, projectsProvider, git, messageBox)
        {
        }

        protected override void OnExecute(object parameter)
        {
            var project = TargetProject(parameter);
            if (project == null)
            {
                return;
            }

            Guarded(() =>
            {
                var url = GitRemoteDialog.Ask(project.Name, Git.RemoteUrl(project));
                if (url == null)
                {
                    return;
                }

                Git.SetRemote(project, url);
                MessageBox.Message(string.Format(CodeExplorerUI.GitClient_RemoteSet, project.Name, url));
            });
        }
    }

    /// <summary>Sends the project's commits to its remote.</summary>
    public class GitPushCommand : GitRemoteCommandBase
    {
        public GitPushCommand(IVbeEvents vbeEvents, IProjectsProvider projectsProvider, GitClientService git, IMessageBox messageBox)
            : base(vbeEvents, projectsProvider, git, messageBox)
        {
        }

        protected override void OnExecute(object parameter)
        {
            var project = TargetProject(parameter);
            if (project == null)
            {
                return;
            }

            Guarded(() =>
            {
                var status = SignedIn(() => Git.Push(project));
                MessageBox.Message(string.Format(CodeExplorerUI.GitClient_Pushed, project.Name, status.Url));
            });
        }
    }

    /// <summary>Brings the remote's commits into the project.</summary>
    public class GitPullCommand : GitRemoteCommandBase
    {
        public GitPullCommand(IVbeEvents vbeEvents, IProjectsProvider projectsProvider, GitClientService git, IMessageBox messageBox)
            : base(vbeEvents, projectsProvider, git, messageBox)
        {
        }

        protected override void OnExecute(object parameter)
        {
            var project = TargetProject(parameter);
            if (project == null)
            {
                return;
            }

            Guarded(() =>
            {
                var outcome = SignedIn(() => WithProgress(progress => Git.Pull(project, progress)));
                if (outcome.Kind == GitMergeKind.Conflicts)
                {
                    MessageBox.NotifyWarn(string.Format(CodeExplorerUI.GitClient_PullConflicts, string.Join(Environment.NewLine, outcome.Conflicts)), CodeExplorerUI.GitClient_DialogTitle);
                    return;
                }

                var lines = new List<string>
                {
                    outcome.Kind == GitMergeKind.UpToDate
                        ? string.Format(CodeExplorerUI.GitClient_PullUpToDate, project.Name)
                        : string.Format(CodeExplorerUI.GitClient_Pulled, project.Name)
                };
                lines.AddRange(outcome.Errors);
                lines.AddRange(outcome.Notes);
                if (outcome.BackupFolder != null)
                {
                    lines.Add(string.Format(CodeExplorerUI.GitSync_Backup, outcome.BackupFolder));
                }

                MessageBox.Message(string.Join(Environment.NewLine + Environment.NewLine, lines));
            });
        }
    }

    /// <summary>Says how far apart the project's history and the remote's are.</summary>
    public class GitRemoteStatusCommand : GitRemoteCommandBase
    {
        public GitRemoteStatusCommand(IVbeEvents vbeEvents, IProjectsProvider projectsProvider, GitClientService git, IMessageBox messageBox)
            : base(vbeEvents, projectsProvider, git, messageBox)
        {
        }

        protected override void OnExecute(object parameter)
        {
            var project = TargetProject(parameter);
            if (project == null)
            {
                return;
            }

            Guarded(() =>
            {
                var status = SignedIn(() => Git.RemoteStatus(project));
                MessageBox.Message(status.Ahead == null
                    ? string.Format(CodeExplorerUI.GitClient_RemoteEmpty, status.Url)
                    : string.Format(CodeExplorerUI.GitClient_RemoteDistance, status.Url, status.Ahead, status.Behind));
            });
        }
    }
}
