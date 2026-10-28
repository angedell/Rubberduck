using System.Collections.Generic;
using System.Diagnostics;
using Rubberduck.PackageResources;
using Rubberduck.VBEditor;

namespace Rubberduck.Navigation.CodeExplorer
{
    /// <summary>
    /// A ribbon XML part or a picture inside the document's file. Ribbon XML and rubberduck.json open in the text editor; pictures can be ticked in the selection mode to be removed in bulk.
    /// </summary>
    [DebuggerDisplay("{" + nameof(Name) + "}")]
    public sealed class CodeExplorerResourceViewModel : CodeExplorerItemViewModelBase
    {
        public CodeExplorerResourceViewModel(ICodeExplorerNode parent, PackageResource resource)
            : base(parent, parent?.Declaration)
        {
            Resource = resource;
        }

        public PackageResource Resource { get; private set; }

        /// <summary>The file name in the package, e.g. customUI14.xml; what the service knows it by.</summary>
        public string ResourceName => Resource.Name;

        public bool IsRibbonXml => Resource.Kind == PackageResourceKind.RibbonXml;

        public bool IsImage => Resource.Kind == PackageResourceKind.Image;

        /// <summary>The document's rubberduck.json.</summary>
        public bool IsConfig => Resource.Kind == PackageResourceKind.Config;

        /// <summary>Whether it opens in the text editor: ribbon XML and the config.</summary>
        public bool IsText => !IsImage;

        /// <summary>The document the resource is stored in.</summary>
        public string DocumentPath => (Parent as CodeExplorerResourceFolderViewModel)?.DocumentPath;

        public override string Name
        {
            get
            {
                switch (Resource.Change)
                {
                    case PackageResourceChange.Added:
                    case PackageResourceChange.Modified:
                        return Resource.Name + " *";
                    case PackageResourceChange.Removed:
                        return Resource.Name + " " + Resources.CodeExplorer.CodeExplorerUI.CodeExplorer_ResourceRemoved;
                    default:
                        return Resource.Name;
                }
            }
        }

        public override string NameWithSignature => $"{Name} ({Size(Resource.Size)})";

        public override string PanelTitle => Resource.Name;

        public override string Description => DocumentPath ?? string.Empty;

        public override string ToolTip => Resource.Change == PackageResourceChange.None
            ? NameWithSignature
            : $"{NameWithSignature} - {Resources.CodeExplorer.CodeExplorerUI.CodeExplorer_ResourcesPending}";

        public override QualifiedSelection? QualifiedSelection => null;

        public override bool IsDimmed
        {
            get => Resource.Change == PackageResourceChange.Removed;
            set { /* Follows the staged change. */ }
        }

        public override bool CanBeMarked => IsImage && Resource.Change != PackageResourceChange.Removed;

        public override bool IsErrorState
        {
            get => false;
            set { /* Resources are not parsed, so they cannot be in an error state. */ }
        }

        public override Comparer<ICodeExplorerNode> SortComparer => CodeExplorerItemComparer.Resource;

        public void Update(PackageResource updated)
        {
            var wasMarkable = CanBeMarked;
            Resource = updated;
            OnNameChanged();
            OnPropertyChanged(nameof(ToolTip));
            OnPropertyChanged(nameof(IsDimmed));
            OnPropertyChanged(nameof(CanBeMarked));

            if (wasMarkable && !CanBeMarked)
            {
                IsMarked = false;
            }
        }

        private static string Size(long bytes)
            => bytes < 1024 ? $"{bytes} B" : bytes < 1024 * 1024 ? $"{bytes / 1024.0:0.#} KB" : $"{bytes / (1024.0 * 1024.0):0.#} MB";
    }
}
