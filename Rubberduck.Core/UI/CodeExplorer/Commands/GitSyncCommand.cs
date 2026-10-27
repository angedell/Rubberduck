using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.UI.CodeExplorer.GitSync;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    /// <summary>
    /// Compares the project with a folder once, shows what is out of sync, and syncs what the user picks.
    /// Asks for the folder the first time; the next time it uses the one it remembers.
    /// </summary>
    public class GitSyncCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerCustomFolderViewModel),
            typeof(CodeExplorerProjectViewModel),
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        private readonly IProjectsProvider _projectsProvider;
        private readonly IFileSystemBrowserFactory _dialogFactory;
        private readonly GitSyncService _gitSyncService;
        private readonly IMessageBox _messageBox;

        public GitSyncCommand(
            IVbeEvents vbeEvents,
            IProjectsProvider projectsProvider,
            IFileSystemBrowserFactory dialogFactory,
            GitSyncService gitSyncService,
            IMessageBox messageBox)
            : base(vbeEvents)
        {
            _projectsProvider = projectsProvider;
            _dialogFactory = dialogFactory;
            _gitSyncService = gitSyncService;
            _messageBox = messageBox;

            AddToCanExecuteEvaluation(project => TargetProject(project) != null);
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        private IVBProject TargetProject(object parameter)
        {
            var declaration = (parameter as CodeExplorerItemViewModel)?.Declaration;
            return declaration != null
                ? _projectsProvider.Project(declaration.ProjectId)
                : null;
        }

        protected override void OnExecute(object parameter)
        {
            var project = TargetProject(parameter);
            if (project == null)
            {
                return;
            }

            var folder = _gitSyncService.RememberedFolder(project) ?? AskForFolder(project, null);
            try
            {
                while (folder != null)
                {
                    var scan = WithProgress(progress => _gitSyncService.Scan(project, folder, progress));
                    if (scan.ConfigCreated)
                    {
                        _messageBox.Message(string.Format(CodeExplorerUI.GitSync_ConfigCreated, scan.Config.Path));
                    }

                    if (scan.Config.Errors.Count > 0)
                    {
                        _messageBox.NotifyWarn(
                            string.Format(CodeExplorerUI.GitSync_ConfigErrors, string.Join(Environment.NewLine, scan.Config.Errors)),
                            CodeExplorerUI.GitSync_DialogTitle);
                    }

                    if (scan.Plan.Items.Count == 0)
                    {
                        _messageBox.Message(string.Format(CodeExplorerUI.GitSync_NothingToDo, project.Name, scan.Plan.InSyncCount, scan.Plan.Ignored.Count));
                        return;
                    }

                    switch (GitSyncReviewDialog.Ask(project.Name, scan, items => _gitSyncService.Ignore(scan, items)))
                    {
                        case ReviewOutcome.ChangeFolder:
                            folder = AskForFolder(project, folder);
                            continue;
                        case ReviewOutcome.Apply:
                            Report(WithProgress(progress => _gitSyncService.Apply(project, scan, scan.Plan.Items, progress)));
                            return;
                        default:
                            return;
                    }
                }
            }
            catch (OperationCanceledException)
            {
                _messageBox.Message(CodeExplorerUI.GitSync_Cancelled);
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "Git sync failed.");
                _messageBox.NotifyWarn(string.Format(CodeExplorerUI.GitSync_Failed, exception.Message), CodeExplorerUI.GitSync_DialogTitle);
            }
        }

        //Scanning and syncing take a while on big projects; the window shows how far they are and lets the user stop.
        private static T WithProgress<T>(Func<ISyncProgress, T> work)
        {
            using (var progress = new GitSyncProgressDialog())
            {
                progress.Show();
                return work(progress);
            }
        }

        private void Report(GitSyncResult result)
        {
            var lines = new List<string> { string.Format(CodeExplorerUI.GitSync_Summary, result.Imported, result.Exported, result.Errors.Count) };
            lines.AddRange(result.Notes);
            lines.AddRange(result.Errors);
            if (result.BackupFolder != null)
            {
                lines.Add(string.Format(CodeExplorerUI.GitSync_Backup, result.BackupFolder));
            }

            _messageBox.Message(string.Join(Environment.NewLine + Environment.NewLine, lines));
        }

        private string AskForFolder(IVBProject project, string initialFolder)
        {
            var description = string.Format(CodeExplorerUI.GitSync_FolderDialogDescription, project.Name);
            using (var browser = _dialogFactory.CreateFolderBrowser(description, true, initialFolder ?? string.Empty))
            {
                return browser.ShowDialog() == DialogResult.OK
                    ? browser.SelectedPath
                    : null;
            }
        }
    }
}
