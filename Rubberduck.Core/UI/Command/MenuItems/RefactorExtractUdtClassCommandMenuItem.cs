using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorExtractUdtClassCommandMenuItem : CommandMenuItemBase
    {
        public RefactorExtractUdtClassCommandMenuItem(RefactorExtractUdtClassCommand command) : base(command)
        {
        }

        public override string Key => "RefactorMenu_ExtractUdtClass";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.ExtractUdtClass;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
