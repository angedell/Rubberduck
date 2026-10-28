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
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    /// <summary>
    /// Copies or moves notes, standard modules and classes to another open project.
    /// It works on one module (the context menu), on every module in a folder (the context menu of a folder), or on the modules ticked in the selection mode.
    /// Forms and document modules are left out: they belong to their host and carry designers, binary parts and bindings that do not travel as text.
    /// A name the target already uses is never overwritten: the user is asked once whether to keep both under numbered names.
    /// </summary>
    public abstract class TransferToProjectCommandBase : ComCommandBase
    {
        private readonly IProjectsProvider _projectsProvider;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IParseManager _parseManager;
        private readonly IMessageBox _messageBox;
        private readonly IFileSystem _fileSystem;

        protected TransferToProjectCommandBase(
            IProjectsProvider projectsProvider,
            IDeclarationFinderProvider declarationFinderProvider,
            IParseManager parseManager,
            IMessageBox messageBox,
            IFileSystem fileSystem,
            IVbeEvents vbeEvents)
            : base(vbeEvents)
        {
            _projectsProvider = projectsProvider;
            _declarationFinderProvider = declarationFinderProvider;
            _parseManager = parseManager;
            _messageBox = messageBox;
            _fileSystem = fileSystem;

            AddToCanExecuteEvaluation(parameter => Modules(parameter).Any() && _projectsProvider.Projects().Count() > 1);
        }

        /// <summary>Raised after the modules have been transferred.</summary>
        public event EventHandler Completed;

        protected abstract bool Moves { get; }

        protected abstract string PickerTitle { get; }

        //The modules the parameter stands for: the node itself, the modules under a folder, or the ticked ones.
        private static IEnumerable<CodeExplorerComponentViewModel> Modules(object parameter)
        {
            switch (parameter)
            {
                case CodeExplorerComponentViewModel module:
                    return CanTravel(module) ? new[] { module } : Enumerable.Empty<CodeExplorerComponentViewModel>();
                case CodeExplorerCustomFolderViewModel folder:
                    return Descendants(folder.Children).Where(CanTravel);
                case IEnumerable<ICodeExplorerNode> nodes:
                    return Descendants(nodes).Where(module => module.IsMarked && module.CanBeMarked && CanTravel(module));
                default:
                    return Enumerable.Empty<CodeExplorerComponentViewModel>();
            }
        }

        private static IEnumerable<CodeExplorerComponentViewModel> Descendants(IEnumerable<ICodeExplorerNode> nodes)
        {
            foreach (var node in nodes)
            {
                if (node is CodeExplorerComponentViewModel module)
                {
                    yield return module;
                    continue;
                }

                foreach (var inner in Descendants(node.Children))
                {
                    yield return inner;
                }
            }
        }

        private static bool CanTravel(CodeExplorerComponentViewModel module)
        {
            var type = module.Declaration?.QualifiedModuleName.ComponentType;
            return type == ComponentType.StandardModule || type == ComponentType.ClassModule;
        }

        protected override void OnExecute(object parameter)
        {
            var modules = Modules(parameter)
                .Select(module => module.Declaration.QualifiedModuleName)
                .Distinct()
                .ToList();
            if (modules.Count == 0)
            {
                return;
            }

            var sourceIds = modules.Select(module => module.ProjectId).Distinct().ToList();
            var targets = _projectsProvider.Projects()
                .Where(project => sourceIds.Count > 1 || project.ProjectId != sourceIds[0])
                .Select(project => new KeyValuePair<string, string>(project.ProjectId, ProjectLabel(project.Project)))
                .ToList();

            var targetId = ProjectPickerDialog.Pick(PickerTitle, targets);
            if (targetId == null)
            {
                return;
            }

            //What is already in the target is left alone, and so are modules that are in it already.
            modules = modules.Where(module => module.ProjectId != targetId).ToList();
            if (modules.Count == 0)
            {
                return;
            }

            Exception failure = null;
            var done = 0;
            var suspension = _parseManager.OnSuspendParser(this, new[] { ParserState.Ready }, () =>
            {
                try
                {
                    done = Transfer(modules, targetId);
                }
                catch (Exception exception) when (exception is System.Runtime.InteropServices.COMException || exception is System.IO.IOException || exception is InvalidOperationException)
                {
                    failure = exception;
                }
            });

            if (failure != null || suspension.Outcome != SuspensionOutcome.Completed)
            {
                var reason = failure?.Message ?? suspension.Outcome.ToString();
                Logger.Warn(failure, "Could not {0} {1} module(s) to project '{2}': {3}", Moves ? "move" : "copy", modules.Count, targetId, reason);
                _messageBox.NotifyWarn(string.Format(CodeExplorerUI.Transfer_Failed, done, modules.Count, reason), PickerTitle);
            }

            if (done > 0)
            {
                Completed?.Invoke(this, EventArgs.Empty);
            }
        }

        private string ProjectLabel(IVBProject project)
        {
            var file = project.FileName;
            return string.IsNullOrEmpty(file)
                ? project.Name + " " + CodeExplorerUI.ProjectPicker_Unsaved
                : project.Name + "  (" + _fileSystem.Path.GetFileName(file) + ")";
        }

        //The number transferred. A move removes a module only once its copy exists.
        private int Transfer(IReadOnlyList<QualifiedModuleName> sources, string targetId)
        {
            var names = ComponentNames(targetId);
            var notesInTarget = NotesIn(targetId);
            string targetName;
            using (var project = _projectsProvider.Project(targetId))
            {
                targetName = project.Name;
            }

            var plans = sources.Select(source => PlanFor(source, names, notesInTarget)).ToList();

            if (!ConfirmClashes(plans, targetName) || (Moves && !ConfirmReferrers(plans.Select(plan => plan.Source).ToList())))
            {
                return 0;
            }

            var done = 0;
            foreach (var plan in plans)
            {
                if (plan.Note)
                {
                    CopyNote(plan, targetId, names);
                }
                else
                {
                    CopyModule(plan, targetId);
                }

                if (Moves)
                {
                    using (var sourceComponents = _projectsProvider.ComponentsCollection(plan.Source.ProjectId))
                    using (var component = _projectsProvider.Component(plan.Source))
                    {
                        sourceComponents.RemoveSafely(component);
                    }
                }

                done++;
            }

            return done;
        }

        private class TransferPlan
        {
            public QualifiedModuleName Source;
            public string Code;
            public bool Note;

            /// <summary>What the module is called in the target: its name, or for a note its display name.</summary>
            public string Name;

            /// <summary>The name asked for, which the target already uses.</summary>
            public string TakenName;
            public string Folder;
        }

        private TransferPlan PlanFor(QualifiedModuleName source, ICollection<string> targetNames, ISet<string> notesInTarget)
        {
            string code;
            using (var component = _projectsProvider.Component(source))
            using (var codeModule = component.CodeModule)
            {
                code = codeModule.CountOfLines == 0 ? string.Empty : codeModule.GetLines(1, codeModule.CountOfLines);
            }

            var plan = new TransferPlan { Source = source, Code = code, Note = NoteModule.IsNote(code), Name = source.ComponentName };
            if (plan.Note)
            {
                plan.Folder = NoteModule.FolderOf(NoteModule.Annotations(code));
                plan.Name = NoteModule.DisplayName(source.ComponentName, code);
                if (notesInTarget.Contains(NoteModule.PathOf(plan.Folder, plan.Name)))
                {
                    plan.TakenName = plan.Name;
                    plan.Name = UniqueNoteName(plan.Name, plan.Folder, notesInTarget);
                }

                notesInTarget.Add(NoteModule.PathOf(plan.Folder, plan.Name));
            }
            else
            {
                if (targetNames.Contains(plan.Name, StringComparer.OrdinalIgnoreCase))
                {
                    plan.TakenName = plan.Name;
                    plan.Name = UniqueName(plan.Name, targetNames);
                }

                targetNames.Add(plan.Name);
            }

            return plan;
        }

        //One question for all the clashes, so that cancelling it leaves everything as it was.
        private bool ConfirmClashes(IEnumerable<TransferPlan> plans, string targetName)
        {
            var clashes = plans.Where(plan => plan.TakenName != null)
                .Select(plan => plan.Note ? NoteModule.PathOf(plan.Folder, plan.TakenName) + " > " + plan.Name : plan.TakenName + " > " + plan.Name)
                .ToList();

            return clashes.Count == 0
                   || _messageBox.ConfirmYesNo(string.Format(CodeExplorerUI.Transfer_NameTaken, targetName, string.Join(Environment.NewLine, clashes)), PickerTitle, true);
        }

        private void CopyNote(TransferPlan plan, string targetId, ICollection<string> targetModuleNames)
        {
            var annotations = NoteModule.Annotations(plan.Code);
            var newCode = NoteModule.ToModuleCode(NoteModule.ToMarkdown(plan.Code), NoteModule.WithMarker(annotations, plan.Name));
            using (var components = _projectsProvider.ComponentsCollection(targetId))
            using (var added = components.Add(ComponentType.StandardModule))
            {
                var moduleName = NoteModule.NewModuleName(targetModuleNames);
                targetModuleNames.Add(moduleName);
                added.Name = moduleName;
                using (var codeModule = added.CodeModule)
                {
                    codeModule.Clear();
                    codeModule.AddFromString(newCode);
                }
            }
        }

        private void CopyModule(TransferPlan plan, string targetId)
        {
            var folder = _fileSystem.Path.Combine(_fileSystem.Path.GetTempPath(), "Rubberduck-transfer-" + Guid.NewGuid().ToString("N"));
            try
            {
                string file;
                using (var component = _projectsProvider.Component(plan.Source))
                {
                    file = component.ExportAsSourceFile(folder);
                }

                using (var components = _projectsProvider.ComponentsCollection(targetId))
                using (var imported = components.Import(file))
                {
                    if (!string.Equals(imported.Name, plan.Name, StringComparison.Ordinal))
                    {
                        imported.Name = plan.Name;
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
        }

        //Modules that use the moved ones and stay where they are lose them. The ones that move together keep each other.
        private bool ConfirmReferrers(IReadOnlyCollection<QualifiedModuleName> moved)
        {
            var finder = _declarationFinderProvider.DeclarationFinder;
            var lines = new List<string>();
            foreach (var source in moved)
            {
                var module = finder.UserDeclarations(DeclarationType.Module)
                    .FirstOrDefault(declaration => declaration.QualifiedName.QualifiedModuleName.Equals(source));
                if (module == null)
                {
                    continue;
                }

                var users = module.References.Concat(finder.Members(source).SelectMany(member => member.References))
                    .Select(reference => reference.QualifiedModuleName)
                    .Where(name => !name.Equals(source) && !moved.Contains(name))
                    .Select(name => name.ProjectName + "." + name.ComponentName)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (users.Count > 0)
                {
                    lines.Add(source.ComponentName + ": " + string.Join(", ", users));
                }
            }

            return lines.Count == 0
                   || _messageBox.ConfirmYesNo(string.Format(CodeExplorerUI.Transfer_Referrers, string.Join(Environment.NewLine, lines)), PickerTitle, false);
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
        public CopyToProjectCommand(IProjectsProvider projectsProvider, IDeclarationFinderProvider declarationFinderProvider, IParseManager parseManager, IMessageBox messageBox, IFileSystem fileSystem, IVbeEvents vbeEvents)
            : base(projectsProvider, declarationFinderProvider, parseManager, messageBox, fileSystem, vbeEvents)
        {
        }

        protected override bool Moves => false;

        protected override string PickerTitle => CodeExplorerUI.ProjectPicker_CopyTitle;
    }

    public class MoveToProjectCommand : TransferToProjectCommandBase
    {
        public MoveToProjectCommand(IProjectsProvider projectsProvider, IDeclarationFinderProvider declarationFinderProvider, IParseManager parseManager, IMessageBox messageBox, IFileSystem fileSystem, IVbeEvents vbeEvents)
            : base(projectsProvider, declarationFinderProvider, parseManager, messageBox, fileSystem, vbeEvents)
        {
        }

        protected override bool Moves => true;

        protected override string PickerTitle => CodeExplorerUI.ProjectPicker_MoveTitle;
    }
}
