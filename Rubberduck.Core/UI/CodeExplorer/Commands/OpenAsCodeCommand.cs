using System;
using System.Collections.Generic;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    /// <summary>Opens a note in a code pane, where it is a module like any other: the way to see the annotations, or to rename it or move it to another folder with the usual commands.</summary>
    public class OpenAsCodeCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes = { typeof(CodeExplorerComponentViewModel) };

        private readonly INavigateCommand _navigate;

        public OpenAsCodeCommand(INavigateCommand navigate, IVbeEvents vbeEvents)
            : base(vbeEvents)
        {
            _navigate = navigate;

            AddToCanExecuteEvaluation(EvaluateCanExecute);
        }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        private static bool EvaluateCanExecute(object parameter)
        {
            return parameter is CodeExplorerComponentViewModel note && note.IsNote && note.QualifiedSelection.HasValue;
        }

        protected override void OnExecute(object parameter)
        {
            if (parameter is CodeExplorerComponentViewModel note && note.QualifiedSelection.HasValue)
            {
                _navigate.Execute(note.QualifiedSelection.Value.GetNavitationArgs());
            }
        }
    }
}
