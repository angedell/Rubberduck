using System;
using System.Collections.Generic;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    internal static class ResourceNodes
    {
        public static readonly Type[] FolderAndResource = { typeof(CodeExplorerResourceFolderViewModel), typeof(CodeExplorerResourceViewModel) };

        public static CodeExplorerResourceFolderViewModel FolderOf(object parameter)
        {
            switch (parameter)
            {
                case CodeExplorerResourceFolderViewModel folder:
                    return folder;
                case CodeExplorerResourceViewModel resource:
                    return resource.Parent as CodeExplorerResourceFolderViewModel;
                default:
                    return null;
            }
        }
    }

    /// <summary>Adds pictures from files to the document's ribbon resources.</summary>
    public class AddResourceImagesCommand : CodeExplorerCommandBase
    {
        private readonly ResourceEditorService _editor;

        public AddResourceImagesCommand(ResourceEditorService editor, IVbeEvents vbeEvents) : base(vbeEvents)
        {
            _editor = editor;

            //A picture is only referred to by ribbon XML, so there has to be some.
            AddToCanExecuteEvaluation(parameter => ResourceNodes.FolderOf(parameter)?.HasRibbon ?? false);
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ResourceNodes.FolderAndResource;

        protected override void OnExecute(object parameter) => _editor.AddImages(ResourceNodes.FolderOf(parameter).DocumentPath);
    }

    /// <summary>Starts a ribbon XML, from a template, in a document that has none.</summary>
    public class AddRibbonXmlCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes = { typeof(CodeExplorerResourceFolderViewModel) };

        private readonly ResourceEditorService _editor;

        public AddRibbonXmlCommand(ResourceEditorService editor, IVbeEvents vbeEvents) : base(vbeEvents)
        {
            _editor = editor;
            AddToCanExecuteEvaluation(parameter => parameter is CodeExplorerResourceFolderViewModel folder && !folder.HasRibbon);
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        protected override void OnExecute(object parameter)
        {
            var folder = (CodeExplorerResourceFolderViewModel)parameter;
            _editor.AddRibbonXml(folder.DocumentPath, OpenCommand.ProjectName(folder));
        }
    }

    /// <summary>Removes the picture that is selected. Several at once go through the selection mode.</summary>
    public class RemoveResourceCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes = { typeof(CodeExplorerResourceViewModel) };

        private readonly ResourceEditorService _editor;

        public RemoveResourceCommand(ResourceEditorService editor, IVbeEvents vbeEvents) : base(vbeEvents)
        {
            _editor = editor;
            AddToCanExecuteEvaluation(parameter => parameter is CodeExplorerResourceViewModel resource && resource.CanBeMarked);
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        protected override void OnExecute(object parameter)
        {
            var resource = (CodeExplorerResourceViewModel)parameter;
            _editor.RemoveImages(resource.DocumentPath, new[] { resource.ResourceName });
        }
    }

    /// <summary>Throws away the ribbon XML and picture changes that are waiting for the document to be closed.</summary>
    public class DiscardResourceChangesCommand : CodeExplorerCommandBase
    {
        private readonly ResourceEditorService _editor;

        public DiscardResourceChangesCommand(ResourceEditorService editor, IVbeEvents vbeEvents) : base(vbeEvents)
        {
            _editor = editor;
            AddToCanExecuteEvaluation(parameter => ResourceNodes.FolderOf(parameter)?.HasPendingChanges ?? false);
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ResourceNodes.FolderAndResource;

        protected override void OnExecute(object parameter) => _editor.Discard(ResourceNodes.FolderOf(parameter).DocumentPath);
    }
}
