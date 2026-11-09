using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.ExtractUdtClass;

namespace Rubberduck.UI.Refactorings
{
    public class ExtractUdtClassViewModel : RefactoringViewModelBase<ExtractUdtClassModel>
    {
        public ExtractUdtClassViewModel(ExtractUdtClassModel model) : base(model)
        {}

        public string ClassName
        {
            get => Model.ClassName;
            set
            {
                Model.ClassName = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidClassName));
                OnPropertyChanged(nameof(CanOk));
            }
        }

        public bool IsValidClassName
            => VBAIdentifierValidator.IsValidIdentifier(ClassName, DeclarationType.ClassModule)
               && !Model.ConflictFinder.IsConflictingModuleName(ClassName);

        /// <summary>What stops the class from being written with the chosen style: member names it needs for itself. Empty if nothing does.</summary>
        public string ClashMessage => string.Join(System.Environment.NewLine, Model.Clashes);

        public bool HasClashes => Model.Clashes.Count > 0;

        public bool CanOk => IsValidClassName && !HasClashes;

        public bool IsPlain
        {
            get => Model.Style == ExtractUdtClassStyle.Plain;
            set
            {
                if (value)
                {
                    SetStyle(ExtractUdtClassStyle.Plain);
                }
            }
        }

        public bool IsFluent
        {
            get => Model.Style == ExtractUdtClassStyle.Fluent;
            set
            {
                if (value)
                {
                    SetStyle(ExtractUdtClassStyle.Fluent);
                }
            }
        }

        private void SetStyle(ExtractUdtClassStyle style)
        {
            Model.Style = style;
            OnPropertyChanged(nameof(IsPlain));
            OnPropertyChanged(nameof(IsFluent));
            OnPropertyChanged(nameof(ClashMessage));
            OnPropertyChanged(nameof(HasClashes));
            OnPropertyChanged(nameof(CanOk));
        }
    }
}
