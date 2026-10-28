using System.Collections.Generic;
using System.Linq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    /// <summary>Adds a note (see <see cref="NoteModule"/>): a standard module of comments with the @Notes annotation, named Notes.md, Notes2.md and so on, in the folder that is selected.</summary>
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

            var taken = ModuleNames(node.Declaration.ProjectId);
            var number = 1;
            string moduleName;
            string displayName;
            do
            {
                displayName = "Notes" + (number == 1 ? string.Empty : number.ToString()) + ".md";
                moduleName = NoteModule.ModuleNameFor(displayName);
                number++;
            }
            while (taken.Contains(moduleName));

            var code = NoteModule.ToModuleCode("# Notes", new[] { "'@Notes(\"" + displayName + "\")" });
            _addComponentService.AddComponent(node, ComponentType, code, moduleName);
        }

        private HashSet<string> ModuleNames(string projectId)
        {
            var names = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            using (var components = _projectsProvider.ComponentsCollection(projectId))
            {
                if (components != null)
                {
                    foreach (var component in components)
                    {
                        using (component)
                        {
                            names.Add(component.Name);
                        }
                    }
                }
            }

            return names;
        }
    }
}
