using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.PackageResources;
using Rubberduck.VBEditor;

namespace Rubberduck.Navigation.CodeExplorer
{
    /// <summary>
    /// The "Resources" node of a project: the ribbon XML and pictures stored in the document's file.
    /// </summary>
    [DebuggerDisplay("{" + nameof(Name) + "}")]
    public sealed class CodeExplorerResourceFolderViewModel : CodeExplorerItemViewModelBase
    {
        public CodeExplorerResourceFolderViewModel(ICodeExplorerNode parent, string documentPath, IReadOnlyList<PackageResource> resources)
            : base(parent, parent?.Declaration)
        {
            DocumentPath = documentPath;
            Synchronize(resources);
        }

        /// <summary>The file the resources are stored in.</summary>
        public string DocumentPath { get; private set; }

        public override string Name => Resources.CodeExplorer.CodeExplorerUI.CodeExplorer_Resources
                                       + (HasPendingChanges ? " *" : string.Empty);

        public override string NameWithSignature => Name;

        public override string PanelTitle => Resources.CodeExplorer.CodeExplorerUI.CodeExplorer_Resources;

        public override string Description => DocumentPath ?? string.Empty;

        public override string ToolTip => HasPendingChanges
            ? Resources.CodeExplorer.CodeExplorerUI.CodeExplorer_ResourcesPending
            : Description;

        public override QualifiedSelection? QualifiedSelection => null;

        public bool HasPendingChanges => Children.OfType<CodeExplorerResourceViewModel>().Any(child => child.Resource.Change != PackageResourceChange.None);

        /// <summary>Whether the document has ribbon XML (once the staged changes are in), which pictures need.</summary>
        public bool HasRibbon => Children.OfType<CodeExplorerResourceViewModel>().Any(child => child.IsRibbonXml && child.Resource.Change != PackageResourceChange.Removed);

        public override bool CanBeMarked => Children.Any(child => child.CanBeMarked);

        public override bool IsErrorState
        {
            get => false;
            set { /* Resources are not parsed, so they cannot be in an error state. */ }
        }

        public override bool Filtered => false;

        public override Comparer<ICodeExplorerNode> SortComparer => CodeExplorerItemComparer.NodeType;

        /// <summary>Brings the children in line with the listing: adds, removes, and updates what changed.</summary>
        public void Synchronize(IReadOnlyList<PackageResource> resources)
        {
            var listed = resources.ToDictionary(resource => resource.Name, StringComparer.OrdinalIgnoreCase);

            foreach (var child in Children.OfType<CodeExplorerResourceViewModel>().ToList())
            {
                if (listed.TryGetValue(child.ResourceName, out var updated))
                {
                    child.Update(updated);
                    listed.Remove(child.ResourceName);
                }
                else
                {
                    RemoveChild(child);
                }
            }

            foreach (var added in listed.Values)
            {
                AddChild(new CodeExplorerResourceViewModel(this, added) { SortOrder = SortOrder });
            }

            OnPropertyChanged(nameof(CanBeMarked));
            OnPropertyChanged(nameof(HasPendingChanges));
            OnPropertyChanged(nameof(Name));
            OnPropertyChanged(nameof(ToolTip));
        }
    }
}
