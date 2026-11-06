using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorSplitParametersCommandMenuItem : CommandMenuItemBase
    {
        public RefactorSplitParametersCommandMenuItem(RefactorSplitParametersCommand command) : base(command)
        {
        }

        public override string Key => "RefactorMenu_SplitParameters";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.SplitParameters;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
