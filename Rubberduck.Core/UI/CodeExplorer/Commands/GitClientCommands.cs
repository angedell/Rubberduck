using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Git;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.UI.CodeExplorer.GitClient;
using Rubberduck.UI.CodeExplorer.GitSync;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    /// <summary>What the built-in git client's commands have in common: they act on the project of the selected node, and only when the client is switched on.</summary>
    public abstract class GitClientCommandBase : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerCustomFolderViewModel),
            typeof(CodeExplorerProjectViewModel),
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        private readonly IProjectsProvider _projectsProvider;
        protected readonly GitClientService Git;
        protected readonly IMessageBox MessageBox;

        protected GitClientCommandBase(IVbeEvents vbeEvents, IProjectsProvider projectsProvider, GitClientService git, IMessageBox messageBox)
            : base(vbeEvents)
        {
            _projectsProvider = projectsProvider;
            Git = git;
            MessageBox = messageBox;

            AddToCanExecuteEvaluation(parameter => git.Enabled && TargetProject(parameter) != null);
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        protected IVBProject TargetProject(object parameter)
        {
            var declaration = (parameter as CodeExplorerItemViewModel)?.Declaration;
            return declaration != null ? _projectsProvider.Project(declaration.ProjectId) : null;
        }

        //Reading the project takes a while on a big one; the window shows how far it is and lets the user stop.
        protected static T WithProgress<T>(Func<ISyncProgress, T> work)
        {
            using (var progress = new GitSyncProgressDialog())
            {
                progress.Show();
                return work(progress);
            }
        }

        /// <summary>Runs the command's work, turning what can go wrong into something the user can read.</summary>
        protected void Guarded(Action work)
        {
            try
            {
                work();
            }
            catch (OperationCanceledException)
            {
                MessageBox.Message(CodeExplorerUI.GitSync_Cancelled);
            }
            catch (GitClientException exception)
            {
                Logger.Warn(exception, "Git client: refused.");
                MessageBox.NotifyWarn(exception.Message, CodeExplorerUI.GitClient_DialogTitle);
            }
            catch (GitException exception)
            {
                Logger.Error(exception, "Git client: git failed.");
                MessageBox.NotifyWarn(string.Format(CodeExplorerUI.GitClient_Failed, exception.Message), CodeExplorerUI.GitClient_DialogTitle);
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "Git client failed.");
                MessageBox.NotifyWarn(string.Format(CodeExplorerUI.GitClient_Failed, exception.Message), CodeExplorerUI.GitClient_DialogTitle);
            }
        }
    }

    /// <summary>Shows what the project changed since the last commit and saves it as a commit.</summary>
    public class GitCommitCommand : GitClientCommandBase
    {
        public GitCommitCommand(IVbeEvents vbeEvents, IProjectsProvider projectsProvider, GitClientService git, IMessageBox messageBox)
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
                bool? continueCopiedHistory = null;
                IReadOnlyList<GitChange> changes;
                while (true)
                {
                    try
                    {
                        var choice = continueCopiedHistory;
                        changes = WithProgress(progress => Git.Changes(project, choice, progress));
                        break;
                    }
                    catch (ProjectCopiedException copied)
                    {
                        continueCopiedHistory = MessageBox.ConfirmYesNo(
                            string.Format(CodeExplorerUI.GitClient_CopyQuestion, project.Name, copied.CopyOfPath),
                            CodeExplorerUI.GitClient_DialogTitle, true);
                    }
                }

                if (changes.Count == 0)
                {
                    MessageBox.Message(string.Format(CodeExplorerUI.GitClient_NothingToCommit, project.Name));
                    return;
                }

                var message = GitCommitDialog.Ask(project.Name, changes);
                if (message == null)
                {
                    return;
                }

                //The work tree holds what the user just saw; reading the project again would commit something else.
                var outcome = Git.Commit(project, message, continueCopiedHistory, refresh: false);
                var lines = new List<string> { string.Format(CodeExplorerUI.GitClient_Committed, outcome.CommitId?.Substring(0, 7) ?? "-", changes.Count) };
                lines.AddRange(outcome.Notes);
                MessageBox.Message(string.Join(Environment.NewLine + Environment.NewLine, lines));
            });
        }
    }

    /// <summary>Puts a module back as it was in an earlier commit.</summary>
    public class GitRestoreCommand : GitClientCommandBase
    {
        private static readonly Type[] ModuleNodes =
        {
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        public GitRestoreCommand(IVbeEvents vbeEvents, IProjectsProvider projectsProvider, GitClientService git, IMessageBox messageBox)
            : base(vbeEvents, projectsProvider, git, messageBox)
        {
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ModuleNodes;

        protected override void OnExecute(object parameter)
        {
            var project = TargetProject(parameter);
            var module = (parameter as CodeExplorerItemViewModel)?.Declaration?.QualifiedName.QualifiedModuleName.ComponentName;
            if (project == null || module == null)
            {
                return;
            }

            Guarded(() =>
            {
                var commits = Git.ModuleHistory(project, module);
                if (commits.Count == 0)
                {
                    MessageBox.Message(string.Format(CodeExplorerUI.GitClient_NoModuleHistory, module));
                    return;
                }

                var commit = GitRestoreDialog.Ask(module, commits);
                if (commit == null)
                {
                    return;
                }

                if (!MessageBox.ConfirmYesNo(
                    string.Format(CodeExplorerUI.GitClient_RestoreConfirm, module, commit.Id.Substring(0, 7), commit.Date.LocalDateTime.ToString("yyyy-MM-dd HH:mm"), commit.Subject),
                    CodeExplorerUI.GitClient_DialogTitle, false))
                {
                    return;
                }

                var result = WithProgress(progress => Git.RestoreModule(project, module, commit, progress));
                var lines = new List<string>
                {
                    result.Errors.Count == 0
                        ? string.Format(CodeExplorerUI.GitClient_Restored, module, commit.Id.Substring(0, 7))
                        : string.Format(CodeExplorerUI.GitClient_Failed, string.Join(Environment.NewLine, result.Errors))
                };
                lines.AddRange(result.Notes);
                if (result.BackupFolder != null)
                {
                    lines.Add(string.Format(CodeExplorerUI.GitSync_Backup, result.BackupFolder));
                }

                MessageBox.Message(string.Join(Environment.NewLine + Environment.NewLine, lines));
            });
        }
    }

    /// <summary>Lists the commits made for the project.</summary>
    public class GitHistoryCommand : GitClientCommandBase
    {
        public GitHistoryCommand(IVbeEvents vbeEvents, IProjectsProvider projectsProvider, GitClientService git, IMessageBox messageBox)
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
                var commits = Git.History(project);
                if (commits.Count == 0)
                {
                    MessageBox.Message(string.Format(CodeExplorerUI.GitClient_NoHistory, project.Name));
                    return;
                }

                GitHistoryDialog.Show(project.Name, commits.ToList());
            });
        }
    }
}
