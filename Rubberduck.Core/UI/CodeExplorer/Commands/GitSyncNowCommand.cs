using System;
using System.Collections.Generic;
using System.Linq;
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
    /// One-click git sync: syncs the project with the folder it is already synced with, without asking anything.
    /// Falls back to the full <see cref="GitSyncCommand"/> if no folder is known for the project yet.
    /// </summary>
    public class GitSyncNowCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerCustomFolderViewModel),
            typeof(CodeExplorerProjectViewModel),
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        private readonly IProjectsProvider _projectsProvider;
        private readonly GitSyncManager _gitSyncManager;
        private readonly GitSyncCommand _gitSyncCommand;
        private readonly IMessageBox _messageBox;

        public GitSyncNowCommand(
            IVbeEvents vbeEvents,
            IProjectsProvider projectsProvider,
            GitSyncManager gitSyncManager,
            GitSyncCommand gitSyncCommand,
            IMessageBox messageBox)
            : base(vbeEvents)
        {
            _projectsProvider = projectsProvider;
            _gitSyncManager = gitSyncManager;
            _gitSyncCommand = gitSyncCommand;
            _messageBox = messageBox;

            AddToCanExecuteEvaluation(parameter => TargetProject(parameter) != null);
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        private IVBProject TargetProject(object parameter)
        {
            var declaration = (parameter as CodeExplorerItemViewModel)?.Declaration;
            return declaration != null
                ? _projectsProvider.Project(declaration.ProjectId)
                : null;
        }

        private bool ConfirmCreateModule(string name)
        {
            return _messageBox.ConfirmYesNo(
                $"'{name}' could not be imported because the project has no module with that name yet.{Environment.NewLine}{Environment.NewLine}Create an empty user form named '{name}' and import into it?",
                CodeExplorerUI.GitSync_DialogTitle);
        }

        protected override void OnExecute(object parameter)
        {
            var project = TargetProject(parameter);
            if (project == null)
            {
                return;
            }

            var folder = _gitSyncManager.ActiveFolder(project) ?? _gitSyncManager.RememberedFolder(project);
            if (folder == null)
            {
                //Nothing to sync with yet, so ask for the folder.
                _gitSyncCommand.Execute(parameter);
                return;
            }

            try
            {
                var result = _gitSyncManager.Start(project, folder, GitSyncInitialDialog.Ask, ConfirmCreateModule);
                if (result.Cancelled)
                {
                    return;
                }

                var messages = result.Notes.Concat(result.Errors).ToList();
                if (messages.Any())
                {
                    _messageBox.NotifyWarn(string.Join(Environment.NewLine, messages), CodeExplorerUI.GitSync_DialogTitle);
                }
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "Git sync failed.");
                _messageBox.NotifyWarn(string.Format(CodeExplorerUI.GitSync_Failed, exception.Message), CodeExplorerUI.GitSync_DialogTitle);
            }
        }
    }
}
