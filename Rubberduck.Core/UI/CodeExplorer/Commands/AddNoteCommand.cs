using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    /// <summary>
    /// Adds a note (see <see cref="NoteModule"/>): a standard module of comments with the @Notes annotation, named Notes.md, Notes2.md and so on within the selected folder.
    /// The module itself gets a generated name; it is the note's name and folder that identify it.
    /// </summary>
    public class AddNoteCommand : AddComponentCommandBase
    {
        private readonly ICodeExplorerAddComponentService _addComponentService;
        private readonly IProjectsProvider _projectsProvider;

        public AddNoteCommand(
            ICodeExplorerAddComponentService addComponentService,
            IVbeEvents vbeEvents,
            IProjectsProvider projectsProvider)
            : base(addComponentService, vbeEvents, projectsProvider)
        {
            _addComponentService = addComponentService;
            _projectsProvider = projectsProvider;
        }

        public override IEnumerable<ProjectType> AllowableProjectTypes => ProjectTypes.All;

        public override ComponentType ComponentType => ComponentType.StandardModule;

        protected override void OnExecute(object parameter)
        {
            if (!(parameter is CodeExplorerItemViewModel node) || node.Declaration == null)
            {
                return;
            }

            var folder = node is CodeExplorerCustomFolderViewModel customFolder
                ? NoteModule.FolderOf(new[] { customFolder.FolderAttribute })
                : node.Declaration.CustomFolder ?? node.Declaration.ProjectName;

            var moduleNames = new List<string>();
            var notesHere = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            ReadProject(node.Declaration.ProjectId, node.Declaration.ProjectName, folder, moduleNames, notesHere);

            var number = 1;
            string displayName;
            do
            {
                displayName = "Notes" + (number == 1 ? string.Empty : number.ToString()) + ".md";
                number++;
            }
            while (notesHere.Contains(displayName));

            var code = NoteModule.ToModuleCode("# Notes", new[] { NoteModule.WithMarker(null, displayName)[0] });
            _addComponentService.AddComponent(node, ComponentType, code, NoteModule.NewModuleName(moduleNames));
        }

        //The names of the project's modules, and the names of the notes already in the folder; the annotations are at the top of a module, so only the top is read.
        private void ReadProject(string projectId, string projectName, string folder, List<string> moduleNames, HashSet<string> notesInFolder)
        {
            using (var components = _projectsProvider.ComponentsCollection(projectId))
            {
                if (components == null)
                {
                    return;
                }

                foreach (var component in components)
                {
                    using (component)
                    {
                        moduleNames.Add(component.Name);
                        using (var codeModule = component.CodeModule)
                        {
                            var count = Math.Min(codeModule.CountOfLines, 30);
                            var head = count == 0 ? string.Empty : codeModule.GetLines(1, count);
                            if (NoteModule.HasMarker(head)
                                && NoteModule.SameFolder(NoteModule.FolderOf(NoteModule.Annotations(head)) ?? projectName, folder))
                            {
                                notesInFolder.Add(NoteModule.DisplayName(component.Name, head));
                            }
                        }
                    }
                }
            }
        }
    }
}
