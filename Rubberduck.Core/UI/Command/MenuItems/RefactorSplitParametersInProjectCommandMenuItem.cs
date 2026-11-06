using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorSplitParametersInProjectCommandMenuItem : CommandMenuItemBase
    {
        public RefactorSplitParametersInProjectCommandMenuItem(RefactorSplitParametersInProjectCommand command) : base(command)
        {
        }

        public override string Key => "RefactorMenu_SplitParametersInProject";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.SplitParametersInProject;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
