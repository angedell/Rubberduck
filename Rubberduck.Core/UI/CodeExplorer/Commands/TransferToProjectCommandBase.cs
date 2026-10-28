using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    /// <summary>
    /// Copies or moves a note, a standard module or a class to another open project.
    /// Forms and document modules are left out: they belong to their host and carry designers, binary parts and bindings that do not travel as text.
    /// A name the target already uses is not overwritten: the user is asked whether to keep both under a new name.
    /// </summary>
    public abstract class TransferToProjectCommandBase : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes = { typeof(CodeExplorerComponentViewModel) };

        private readonly IProjectsProvider _projectsProvider;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IMessageBox _messageBox;
        private readonly IFileSystem _fileSystem;

        protected TransferToProjectCommandBase(IProjectsProvider projectsProvider, IDeclarationFinderProvider declarationFinderProvider, IMessageBox messageBox, IFileSystem fileSystem, IVbeEvents vbeEvents)
            : base(vbeEvents)
        {
            _fileSystem = fileSystem;
            _projectsProvider = projectsProvider;
            _declarationFinderProvider = declarationFinderProvider;
            _messageBox = messageBox;

            AddToCanExecuteEvaluation(EvaluateCanExecute);
        }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        protected abstract bool Moves { get; }

        protected abstract string PickerTitle { get; }

        private bool EvaluateCanExecute(object parameter)
        {
            if (!(parameter is CodeExplorerComponentViewModel node) || node.Declaration == null)
            {
                return false;
            }

            var type = node.Declaration.QualifiedName.QualifiedModuleName.ComponentType;
            return (type == ComponentType.StandardModule || type == ComponentType.ClassModule)
                   && _projectsProvider.Projects().Count() > 1;
        }

        protected override void OnExecute(object parameter)
        {
            if (!(parameter is CodeExplorerComponentViewModel node) || node.Declaration == null)
            {
                return;
            }

            var source = node.Declaration.QualifiedName.QualifiedModuleName;
            var targets = _projectsProvider.Projects()
                .Where(project => project.ProjectId != source.ProjectId)
                .Select(project => new KeyValuePair<string, string>(project.ProjectId, ProjectLabel(project.Project)))
                .ToList();

            var targetId = ProjectPickerDialog.Pick(PickerTitle, targets);
            if (targetId == null)
            {
                return;
            }

            try
            {
                Transfer(source, node.IsNote, targetId);
            }
            catch (Exception exception) when (exception is System.Runtime.InteropServices.COMException || exception is System.IO.IOException || exception is InvalidOperationException)
            {
                Logger.Warn(exception, "Could not {0} '{1}' to project '{2}'.", Moves ? "move" : "copy", source.ComponentName, targetId);
                _messageBox.NotifyWarn(string.Format(CodeExplorerUI.Transfer_Failed, exception.Message), PickerTitle);
            }
        }

        private string ProjectLabel(IVBProject project)
        {
            var file = project.FileName;
            return string.IsNullOrEmpty(file)
                ? project.Name + " " + CodeExplorerUI.ProjectPicker_Unsaved
                : project.Name + "  (" + _fileSystem.Path.GetFileName(file) + ")";
        }

        private void Transfer(QualifiedModuleName source, bool isNote, string targetId)
        {
            var targetNames = ComponentNames(targetId);
            string targetName;
            using (var targetProject = _projectsProvider.Project(targetId))
            {
                targetName = targetProject.Name;
            }

            if (Moves && !ConfirmReferrers(source))
            {
                return;
            }

            string code;
            using (var component = _projectsProvider.Component(source))
            using (var codeModule = component.CodeModule)
            {
                code = codeModule.CountOfLines == 0 ? string.Empty : codeModule.GetLines(1, codeModule.CountOfLines);
            }

            if (isNote && NoteModule.IsNote(code))
            {
                if (!CopyNote(code, targetId, targetName, targetNames))
                {
                    return;
                }
            }
            else if (!CopyModule(source, targetId, targetName, targetNames))
            {
                return;
            }

            if (Moves)
            {
                using (var sourceComponents = _projectsProvider.ComponentsCollection(source.ProjectId))
                using (var component = _projectsProvider.Component(source))
                {
                    sourceComponents.RemoveSafely(component);
                }
            }
        }

        private bool CopyNote(string code, string targetId, string targetName, ICollection<string> targetModuleNames)
        {
            var annotations = NoteModule.Annotations(code);
            var folder = NoteModule.FolderOf(annotations);
            var name = NoteModule.DisplayName(string.Empty, code);
            var taken = NotesIn(targetId);

            var chosen = name;
            if (taken.Contains(NoteModule.PathOf(folder, name)))
            {
                chosen = UniqueNoteName(name, folder, taken);
                if (!_messageBox.ConfirmYesNo(string.Format(CodeExplorerUI.Transfer_NoteTaken, targetName, NoteModule.PathOf(folder, name), chosen), PickerTitle, true))
                {
                    return false;
                }
            }

            var newCode = NoteModule.ToModuleCode(NoteModule.ToMarkdown(code), NoteModule.WithMarker(annotations, chosen));
            using (var components = _projectsProvider.ComponentsCollection(targetId))
            using (var added = components.Add(ComponentType.StandardModule))
            {
                added.Name = NoteModule.NewModuleName(targetModuleNames);
                using (var codeModule = added.CodeModule)
                {
                    codeModule.Clear();
                    codeModule.AddFromString(newCode);
                }
            }

            return true;
        }

        private bool CopyModule(QualifiedModuleName source, string targetId, string targetName, ICollection<string> targetNames)
        {
            var wanted = source.ComponentName;
            if (targetNames.Contains(wanted, StringComparer.OrdinalIgnoreCase))
            {
                wanted = UniqueName(wanted, targetNames);
                if (!_messageBox.ConfirmYesNo(string.Format(CodeExplorerUI.Transfer_NameTaken, targetName, source.ComponentName, wanted), PickerTitle, true))
                {
                    return false;
                }
            }

            var folder = _fileSystem.Path.Combine(_fileSystem.Path.GetTempPath(), "Rubberduck-transfer-" + Guid.NewGuid().ToString("N"));
            try
            {
                string file;
                using (var component = _projectsProvider.Component(source))
                {
                    file = component.ExportAsSourceFile(folder);
                }

                using (var components = _projectsProvider.ComponentsCollection(targetId))
                using (var imported = components.Import(file))
                {
                    if (!string.Equals(imported.Name, wanted, StringComparison.Ordinal))
                    {
                        imported.Name = wanted;
                    }
                }
            }
            finally
            {
                try
                {
                    _fileSystem.Directory.Delete(folder, true);
                }
                catch (System.IO.IOException exception)
                {
                    Logger.Debug(exception, "Could not remove the temporary folder '{0}'.", folder);
                }
            }

            return true;
        }

        private bool ConfirmReferrers(QualifiedModuleName source)
        {
            var finder = _declarationFinderProvider.DeclarationFinder;
            var module = finder.UserDeclarations(DeclarationType.Module)
                .FirstOrDefault(declaration => declaration.QualifiedName.QualifiedModuleName.Equals(source));
            if (module == null)
            {
                return true;
            }

            var users = module.References.Concat(finder.Members(source).SelectMany(member => member.References))
                .Select(reference => reference.QualifiedModuleName)
                .Where(name => !name.Equals(source))
                .Select(name => name.ProjectName + "." + name.ComponentName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return users.Count == 0
                   || _messageBox.ConfirmYesNo(string.Format(CodeExplorerUI.Transfer_Referrers, source.ComponentName, string.Join(", ", users)), PickerTitle, false);
        }

        private ICollection<string> ComponentNames(string projectId)
        {
            var names = new List<string>();
            using (var components = _projectsProvider.ComponentsCollection(projectId))
            {
                foreach (var component in components)
                {
                    using (component)
                    {
                        names.Add(component.Name);
                    }
                }
            }

            return names;
        }

        //The paths (Api/Core/Notes.md) of the notes a project has; the annotations are at the top of a module, so only the top is read.
        private ISet<string> NotesIn(string projectId)
        {
            var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var components = _projectsProvider.ComponentsCollection(projectId))
            {
                foreach (var component in components)
                {
                    using (component)
                    using (var codeModule = component.CodeModule)
                    {
                        var count = Math.Min(codeModule.CountOfLines, 30);
                        var head = count == 0 ? string.Empty : codeModule.GetLines(1, count);
                        if (NoteModule.HasMarker(head))
                        {
                            paths.Add(NoteModule.PathOf(NoteModule.FolderOf(NoteModule.Annotations(head)), NoteModule.DisplayName(component.Name, head)));
                        }
                    }
                }
            }

            return paths;
        }

        /// <summary>A name not in <paramref name="taken"/>: the name with a number after it.</summary>
        internal static string UniqueName(string name, IEnumerable<string> taken)
        {
            var all = new HashSet<string>(taken, StringComparer.OrdinalIgnoreCase);
            var number = 2;
            while (all.Contains(name + number))
            {
                number++;
            }

            return name + number;
        }

        /// <summary>A note name not used in the folder: Notes.md becomes Notes2.md.</summary>
        internal static string UniqueNoteName(string name, string folder, ISet<string> takenPaths)
        {
            var stem = name.Substring(0, name.Length - ".md".Length);
            var number = 2;
            while (takenPaths.Contains(NoteModule.PathOf(folder, stem + number + ".md")))
            {
                number++;
            }

            return stem + number + ".md";
        }
    }

    public class CopyToProjectCommand : TransferToProjectCommandBase
    {
        public CopyToProjectCommand(IProjectsProvider projectsProvider, IDeclarationFinderProvider declarationFinderProvider, IMessageBox messageBox, IFileSystem fileSystem, IVbeEvents vbeEvents)
            : base(projectsProvider, declarationFinderProvider, messageBox, fileSystem, vbeEvents)
        {
        }

        protected override bool Moves => false;

        protected override string PickerTitle => CodeExplorerUI.ProjectPicker_CopyTitle;
    }

    public class MoveToProjectCommand : TransferToProjectCommandBase
    {
        public MoveToProjectCommand(IProjectsProvider projectsProvider, IDeclarationFinderProvider declarationFinderProvider, IMessageBox messageBox, IFileSystem fileSystem, IVbeEvents vbeEvents)
            : base(projectsProvider, declarationFinderProvider, messageBox, fileSystem, vbeEvents)
        {
        }

        protected override bool Moves => true;

        protected override string PickerTitle => CodeExplorerUI.ProjectPicker_MoveTitle;
    }
}
