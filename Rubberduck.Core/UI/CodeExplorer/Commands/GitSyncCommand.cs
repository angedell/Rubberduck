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
    /// Asks for a project folder and keeps it in sync with the project, subfolders included.
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
        private readonly GitSyncManager _gitSyncManager;
        private readonly IMessageBox _messageBox;

        public GitSyncCommand(
            IVbeEvents vbeEvents,
            IProjectsProvider projectsProvider,
            IFileSystemBrowserFactory dialogFactory,
            GitSyncManager gitSyncManager,
            IMessageBox messageBox)
            : base(vbeEvents)
        {
            _projectsProvider = projectsProvider;
            _dialogFactory = dialogFactory;
            _gitSyncManager = gitSyncManager;
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

            var activeFolder = _gitSyncManager.ActiveFolder(project);
            if (activeFolder != null)
            {
                var message = string.Format(CodeExplorerUI.GitSync_AlreadySyncing, project.Name, activeFolder);
                switch (_messageBox.Confirm(message, CodeExplorerUI.GitSync_DialogTitle, ConfirmationOutcome.Cancel))
                {
                    case ConfirmationOutcome.No:
                        _gitSyncManager.Stop(project);
                        return;
                    case ConfirmationOutcome.Cancel:
                        return;
                }
            }

            var folder = AskForFolder(project, activeFolder ?? _gitSyncManager.RememberedFolder(project));
            if (folder == null)
            {
                return;
            }

            try
            {
                var result = _gitSyncManager.Start(project, folder);
                _messageBox.Message(string.Format(
                    CodeExplorerUI.GitSync_Summary,
                    folder, result.Imported, result.Exported, result.LineEndingsFixed, result.FolderAnnotationsFixed, result.Errors.Count)
                    + (result.Errors.Count > 0 ? Environment.NewLine + Environment.NewLine + string.Join(Environment.NewLine, result.Errors) : string.Empty));
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "Git sync could not be started.");
                _messageBox.NotifyWarn(string.Format(CodeExplorerUI.GitSync_Failed, exception.Message), CodeExplorerUI.GitSync_DialogTitle);
            }
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
