using System.Collections.ObjectModel;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.ExtractUdtClass;

namespace Rubberduck.UI.Refactorings
{
    /// <summary>One member of the type in the list of members that Create asks for.</summary>
    public class RequiredMemberItem : ViewModelBase
    {
        private readonly UdtMemberInfo _member;
        private bool _createExists;

        public RequiredMemberItem(UdtMemberInfo member, bool createExists)
        {
            _member = member;
            _createExists = createExists;
        }

        public string Name => _member.Name;
        public string TypeName => _member.IsArray ? _member.TypeName + "()" : _member.TypeName;

        public bool IsRequired
        {
            get => _member.IsRequired;
            set
            {
                _member.IsRequired = value;
                OnPropertyChanged();
            }
        }

        /// <summary>Only the fluent style has a Create to take the arguments, and not every member can be one.</summary>
        public bool CanChange => _createExists && _member.CanBeRequired;

        public void CreateExists(bool exists)
        {
            _createExists = exists;
            OnPropertyChanged(nameof(CanChange));
        }
    }

    public class ExtractUdtClassViewModel : RefactoringViewModelBase<ExtractUdtClassModel>
    {
        public ExtractUdtClassViewModel(ExtractUdtClassModel model) : base(model)
        {
            Members = new ObservableCollection<RequiredMemberItem>(
                model.Members.Select(member => new RequiredMemberItem(member, model.Style == ExtractUdtClassStyle.Fluent)));
        }

        public ObservableCollection<RequiredMemberItem> Members { get; }

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
            foreach (var item in Members)
            {
                item.CreateExists(style == ExtractUdtClassStyle.Fluent);
            }
            OnPropertyChanged(nameof(ClashMessage));
            OnPropertyChanged(nameof(HasClashes));
            OnPropertyChanged(nameof(CanOk));
        }
    }
}
