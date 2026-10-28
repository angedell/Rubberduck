using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources.Ship;
using Rubberduck.Ship;
using Rubberduck.UI.CodeExplorer.Ship;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    /// <summary>
    /// Opens the Ship dialog for a project: writes a clean copy of the saved workbook as an .xlsm or .xlam.
    /// Offered only where <see cref="IShipService.CanShip"/> says so (an Excel project with a file).
    /// </summary>
    public class ShipCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes = { typeof(CodeExplorerProjectViewModel) };

        private readonly IProjectsProvider _projectsProvider;
        private readonly IShipService _shipService;
        private readonly IFileSystemBrowserFactory _dialogFactory;
        private readonly IMessageBox _messageBox;

        public ShipCommand(
            IVbeEvents vbeEvents,
            IProjectsProvider projectsProvider,
            IShipService shipService,
            IFileSystemBrowserFactory dialogFactory,
            IMessageBox messageBox)
            : base(vbeEvents)
        {
            _projectsProvider = projectsProvider;
            _shipService = shipService;
            _dialogFactory = dialogFactory;
            _messageBox = messageBox;

            AddToCanExecuteEvaluation(parameter =>
            {
                var project = TargetProject(parameter);
                return project != null && _shipService.CanShip(project);
            });
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

            try
            {
                var model = new ShipViewModel(_shipService, project, project.Name);
                model.Load();
                ShipDialog.Open(null, model, start => AskForFolder(project.Name, start));
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "Ship failed.");
                _messageBox.NotifyWarn(string.Format(ShipUI.Ship_Done_Failed, exception.Message), string.Format(ShipUI.Ship_Title, project.Name));
            }
        }

        private string AskForFolder(string projectName, string initialFolder)
        {
            using (var browser = _dialogFactory.CreateFolderBrowser(string.Format(ShipUI.Ship_FolderBrowseDescription, projectName), true, initialFolder ?? string.Empty))
            {
                return browser.ShowDialog() == DialogResult.OK
                    ? browser.SelectedPath
                    : null;
            }
        }
    }
}
