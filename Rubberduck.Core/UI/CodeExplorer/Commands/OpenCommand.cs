using System;
using System.Collections.Generic;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class OpenCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerProjectViewModel),
            typeof(CodeExplorerCustomFolderViewModel),
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel),
            typeof(CodeExplorerResourceViewModel)
        };

        private readonly INavigateCommand _openCommand;
        private readonly ResourceEditorService _resourceEditor;

        public OpenCommand(
            INavigateCommand openCommand, 
            ResourceEditorService resourceEditor,
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _openCommand = openCommand;
            _resourceEditor = resourceEditor;

            AddToCanExecuteEvaluation(EvaluateCanExecute);
        }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        private bool EvaluateCanExecute(object parameter)
        {
            //Ribbon XML opens in the XML editor; pictures have nothing to open.
            if (parameter is CodeExplorerResourceViewModel resource)
            {
                return resource.IsText && resource.DocumentPath != null;
            }

            return (parameter as CodeExplorerItemViewModel)?.QualifiedSelection.HasValue ?? false;
        }

        protected override void OnExecute(object parameter)
        {
            if (parameter is CodeExplorerResourceViewModel resource)
            {
                _resourceEditor.EditText(resource.DocumentPath, resource.ResourceName, ProjectName(resource));
                return;
            }

            if (parameter is CodeExplorerItemViewModel item && item.QualifiedSelection.HasValue)
            {
                _openCommand.Execute(item.QualifiedSelection.Value.GetNavitationArgs());
            }
        }

        internal static string ProjectName(ICodeExplorerNode node)
        {
            while (node.Parent != null)
            {
                node = node.Parent;
            }

            return node.Declaration?.IdentifierName ?? string.Empty;
        }
    }
}
