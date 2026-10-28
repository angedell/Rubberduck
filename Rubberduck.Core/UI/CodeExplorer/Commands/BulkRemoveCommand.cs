using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using Newtonsoft.Json;
using System.Windows.Forms;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.CodeExplorer.GitSync;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    /// <summary>
    /// Removes all modules ticked in the code explorer's selection mode in one go,
    /// after offering to export them first, either keeping their folder structure or into a single folder.
    /// Ticked ribbon pictures are removed along with them (staged, until the document is closed).
    /// </summary>
    public class BulkRemoveCommand : ComCommandBase
    {
        private readonly RemoveCommand _removeCommand;
        private readonly IProjectsProvider _projectsProvider;
        private readonly IParseManager _parseManager;
        private readonly IFileSystemBrowserFactory _dialogFactory;
        private readonly IFileSystem _fileSystem;
        private readonly IMessageBox _messageBox;
        private readonly ResourceEditorService _resourceEditor;

        public BulkRemoveCommand(
            RemoveCommand removeCommand,
            IProjectsProvider projectsProvider,
            IParseManager parseManager,
            IFileSystemBrowserFactory dialogFactory,
            IFileSystem fileSystem,
            IMessageBox messageBox,
            ResourceEditorService resourceEditor,
            IVbeEvents vbeEvents)
            : base(vbeEvents)
        {
            _removeCommand = removeCommand;
            _projectsProvider = projectsProvider;
            _parseManager = parseManager;
            _dialogFactory = dialogFactory;
            _fileSystem = fileSystem;
            _messageBox = messageBox;
            _resourceEditor = resourceEditor;

            AddToCanExecuteEvaluation(parameter => MarkedModules(parameter).Any() || MarkedImages(parameter).Any());
        }

        private static IEnumerable<CodeExplorerResourceViewModel> MarkedImages(object parameter)
        {
            return parameter is IEnumerable<ICodeExplorerNode> nodes
                ? MarkedImages(nodes)
                : Enumerable.Empty<CodeExplorerResourceViewModel>();
        }

        private static IEnumerable<CodeExplorerResourceViewModel> MarkedImages(IEnumerable<ICodeExplorerNode> nodes)
        {
            foreach (var node in nodes)
            {
                if (node is CodeExplorerResourceViewModel image)
                {
                    if (image.IsMarked && image.CanBeMarked)
                    {
                        yield return image;
                    }

                    continue;
                }

                foreach (var marked in MarkedImages(node.Children))
                {
                    yield return marked;
                }
            }
        }

        private static IEnumerable<CodeExplorerComponentViewModel> MarkedModules(object parameter)
        {
            return parameter is IEnumerable<ICodeExplorerNode> nodes
                ? MarkedModules(nodes)
                : Enumerable.Empty<CodeExplorerComponentViewModel>();
        }

        private static IEnumerable<CodeExplorerComponentViewModel> MarkedModules(IEnumerable<ICodeExplorerNode> nodes)
        {
            foreach (var node in nodes)
            {
                if (node is CodeExplorerComponentViewModel module)
                {
                    if (module.IsMarked && module.CanBeMarked)
                    {
                        yield return module;
                    }

                    continue;
                }

                foreach (var marked in MarkedModules(node.Children))
                {
                    yield return marked;
                }
            }
        }

        /// <summary>
        /// Raised after the modules have been removed.
        /// </summary>
        public event EventHandler Completed;

        private string SettingsFile => _fileSystem.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck", "bulkremove.json");

        protected override void OnExecute(object parameter)
        {
            var modules = MarkedModules(parameter)
                .GroupBy(module => module.Declaration.QualifiedModuleName)
                .Select(group => group.First())
                .ToList();
            var images = MarkedImages(parameter)
                .GroupBy(image => image.DocumentPath)
                .Where(group => group.Key != null)
                .ToList();
            if (modules.Count == 0 && images.Count == 0)
            {
                return;
            }

            //One question for all the pictures, up front, so cancelling it (or the modules' dialog) leaves everything as it was.
            var pictureCount = images.Sum(group => group.Count());
            if (pictureCount > 0 && !_messageBox.ConfirmYesNo(string.Format(CodeExplorerUI.Resources_RemoveConfirm, pictureCount), CodeExplorerUI.Resources_Caption, false))
            {
                return;
            }

            if (modules.Count > 0 && !RemoveModules(modules))
            {
                return;
            }

            foreach (var document in images)
            {
                _resourceEditor.RemoveImages(document.Key, document.Select(image => image.ResourceName).ToList(), false);
            }

            Completed?.Invoke(this, EventArgs.Empty);
        }

        //False if the user cancelled.
        private bool RemoveModules(List<CodeExplorerComponentViewModel> modules)
        {
            var options = ReadRememberedOptions();
            using (var dialog = new BulkRemoveDialog(modules.Count, options, BrowseForFolder))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return false;
                }

                options = dialog.Result;
            }

            if (options.Export)
            {
                RememberOptions(options);
                if (!TryExport(modules, options))
                {
                    return false;
                }
            }
            else
            {
                //Keep the remembered folder, only remember that the user did not want to export this time.
                var remembered = ReadRememberedOptions();
                remembered.Export = false;
                RememberOptions(remembered);
            }

            Remove(modules);
            return true;
        }

        private string BrowseForFolder(string initialFolder)
        {
            using (var browser = _dialogFactory.CreateFolderBrowser(CodeExplorerUI.BulkRemove_BrowseDescription, true, _fileSystem.Directory.Exists(initialFolder) ? initialFolder : string.Empty))
            {
                return browser.ShowDialog() == DialogResult.OK
                    ? browser.SelectedPath
                    : null;
            }
        }

        private BulkRemoveOptions ReadRememberedOptions()
        {
            try
            {
                if (_fileSystem.File.Exists(SettingsFile))
                {
                    var options = JsonConvert.DeserializeObject<BulkRemoveOptions>(_fileSystem.File.ReadAllText(SettingsFile));
                    if (options != null)
                    {
                        return options;
                    }
                }
            }
            catch (Exception exception)
            {
                Logger.Warn(exception, "Could not read the remembered bulk removal options.");
            }

            return new BulkRemoveOptions();
        }

        private void RememberOptions(BulkRemoveOptions options)
        {
            try
            {
                _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(SettingsFile));
                _fileSystem.File.WriteAllText(SettingsFile, JsonConvert.SerializeObject(options, Formatting.Indented));
            }
            catch (Exception exception)
            {
                Logger.Warn(exception, "Could not remember the bulk removal options.");
            }
        }

        private bool TryExport(IReadOnlyCollection<CodeExplorerComponentViewModel> modules, BulkRemoveOptions options)
        {
            var root = options.ExportFolder;
            var preserveStructure = options.PreserveFolderStructure;

            //Module names are only unique within a project.
            var severalProjects = modules.Select(module => module.Declaration.ProjectId).Distinct().Count() > 1;

            try
            {
                foreach (var module in modules)
                {
                    var qualifiedModuleName = module.Declaration.QualifiedModuleName;
                    var component = _projectsProvider.Component(qualifiedModuleName);
                    if (component == null)
                    {
                        continue;
                    }

                    var folder = root;
                    if (severalProjects)
                    {
                        folder = _fileSystem.Path.Combine(folder, GitSyncText.RelativeDirectory(qualifiedModuleName.ProjectName));
                    }

                    if (preserveStructure)
                    {
                        folder = _fileSystem.Path.Combine(folder, ModuleFolder(module));
                    }

                    _fileSystem.Directory.CreateDirectory(folder);
                    component.ExportAsSourceFile(_fileSystem.Path.Combine(folder, component.Name + component.Type.FileExtension()));
                }
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "Export before bulk removal failed.");
                _messageBox.NotifyWarn(string.Format(CodeExplorerUI.BulkRemove_ExportFailed, exception.Message), CodeExplorerUI.BulkRemove_Caption);
                return false;
            }

            return true;
        }

        //Modules without a @Folder annotation default to the project name, which is the root of the tree.
        private static string ModuleFolder(CodeExplorerComponentViewModel module)
        {
            var folder = module.Declaration.CustomFolder;
            return folder == module.Declaration.QualifiedModuleName.ProjectName
                ? string.Empty
                : GitSyncText.RelativeDirectory(folder);
        }

        private void Remove(IEnumerable<CodeExplorerComponentViewModel> modules)
        {
            var qualifiedModuleNames = modules.Select(module => module.Declaration.QualifiedModuleName).ToList();

            var suspension = _parseManager.OnSuspendParser(this, new[] { ParserState.Ready }, () =>
            {
                foreach (var qualifiedModuleName in qualifiedModuleNames)
                {
                    //The modules were exported above, if the user wanted that, so don't ask again.
                    _removeCommand.RemoveComponent(qualifiedModuleName, false);
                }
            });

            if (suspension.Outcome != SuspensionOutcome.Completed)
            {
                Logger.Warn("Bulk removal failed due to suspension failure: {0}", suspension.Outcome);
                _messageBox.NotifyWarn(string.Format(CodeExplorerUI.BulkRemove_SuspendFailed, suspension.Outcome), CodeExplorerUI.BulkRemove_Caption);
            }
        }
    }
}
