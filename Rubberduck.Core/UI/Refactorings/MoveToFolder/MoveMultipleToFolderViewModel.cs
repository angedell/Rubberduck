using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.Refactorings;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources;
using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.CodeAnalysis;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.CodeExplorer;

namespace Rubberduck.UI.Refactorings.MoveToFolder
{
    public class MoveMultipleToFolderViewModel : RefactoringViewModelBase<MoveMultipleToFolderModel>, IFolderBrowsing
    {
        private readonly IDeclarationFinderProvider _finder;

        public MoveMultipleToFolderViewModel(MoveMultipleToFolderModel model) 
            : base(model)
        {}

        public MoveMultipleToFolderViewModel(MoveMultipleToFolderModel model, IDeclarationFinderProvider finder)
            : base(model)
        {
            _finder = finder;
        }

        private ICollection<ModuleDeclaration> Targets => Model.Targets;

        public string Instructions
        {
            get
            {
                if (Targets == null || !Targets.Any())
                {
                    return string.Empty;
                }

                if (Targets.Count == 1)
                {
                    var target = Targets.First();
                    var moduleName = target.IdentifierName;
                    var declarationType = CodeAnalysisUI.ResourceManager.GetString("DeclarationType_" + target.DeclarationType, CultureInfo.CurrentUICulture);
                    var currentFolder = target.CustomFolder;
                    return string.Format(RefactoringsUI.MoveToFolderDialog_InstructionsLabelText, declarationType, moduleName, currentFolder);
                }

                return string.Format(RefactoringsUI.MoveMultipleToFolderDialog_InstructionsLabelText);
            }
        }

        //A folder belongs to a project: the folders of the others are not places to move to.
        private HashSet<string> ProjectIds => new HashSet<string>(Targets?.Select(target => target.ProjectId) ?? Enumerable.Empty<string>());

        /// <summary>The folders of the project the modules are in, for the folder picker.</summary>
        public IReadOnlyList<string> KnownFolders => (_finder?.DeclarationFinder
            .UserDeclarations(DeclarationType.Module)
            .OfType<ModuleDeclaration>()
            .Where(module => ProjectIds.Contains(module.ProjectId))
            .Select(module => module.CustomFolder)
            .Where(folder => !string.IsNullOrWhiteSpace(folder))
            .Distinct(System.StringComparer.OrdinalIgnoreCase)
            .OrderBy(folder => folder, System.StringComparer.OrdinalIgnoreCase)
            .ToList() ?? new List<string>());

        public string NewFolder
        {
            get => Model.TargetFolder;
            set
            {
                Model.TargetFolder = value;
                ValidateFolder();
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidFolder));
            }
        }

        private void ValidateFolder()
        {
            var errors = new List<string>();
            
            if (string.IsNullOrEmpty(NewFolder))
            {
                errors.Add(RefactoringsUI.MoveFolders_EmptyFolderName);
            }
            else
            {
                if (NewFolder.Any(char.IsControl))
                {
                    errors.Add(RefactoringsUI.MoveFolders_ControlCharacter);
                }

                if (NewFolder.Split(FolderExtensions.FolderDelimiter).Any(string.IsNullOrEmpty))
                {
                    errors.Add(RefactoringsUI.MoveFolders_EmptySubfolderName);
                }
            }

            if (errors.Any())
            {
                SetErrors(nameof(NewFolder), errors);
            }
            else
            {
                ClearErrors();
            }
        }

        public bool IsValidFolder => Targets != null 
                                     && Targets.Any()
                                     && !HasErrors;

        protected override void DialogOk()
        {
            if (Targets == null || !Targets.Any())
            {
                base.DialogCancel();
            }
            else
            {
                base.DialogOk();
            }
        }
    }
}
