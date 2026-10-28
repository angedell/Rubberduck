using System.IO.Abstractions;
using System.Windows.Forms;
using NLog;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>Opens a note (see <see cref="NoteModule"/>) in the markdown editor and writes what the user saves back into the module.</summary>
    public class NoteEditorService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IProjectsProvider _projects;
        private readonly IMessageBox _messageBox;
        private readonly IFileSystem _fileSystem;

        public NoteEditorService(IProjectsProvider projects, IMessageBox messageBox, IFileSystem fileSystem)
        {
            _projects = projects;
            _messageBox = messageBox;
            _fileSystem = fileSystem;
        }

        /// <summary>Opens the note in the markdown editor. False if the module is not a note after all (it has code in it): nothing is opened, and the caller should open it as code.</summary>
        public bool Edit(QualifiedModuleName module, string projectName)
        {
            string opened;
            using (var component = _projects.Component(module))
            using (var codeModule = component.CodeModule)
            {
                opened = codeModule.Content();
            }

            //The last word on what is a note is the text: saving from the markdown editor replaces everything in the module.
            if (!NoteModule.IsNote(opened))
            {
                Logger.Info("'{0}' in project '{1}' has the @Notes marker in the tree but is not a note by its text (code in it, or no marker); opening it as code.", module.ComponentName, projectName);
                return false;
            }

            var displayName = NoteModule.DisplayName(module.ComponentName, opened);
            Logger.Debug("Opening note '{0}' (module '{1}', project '{2}', {3} characters).", displayName, module.ComponentName, projectName, opened.Length);

            var annotations = NoteModule.Annotations(opened);
            var layoutKey = NoteEditorLayouts.KeyFor(ProjectFile(module), projectName, module.ComponentName);
            var layouts = NoteEditorLayouts.Load(_fileSystem);
            var preferences = ResourceEditorPreferences.Load(_fileSystem);

            string saved;
            string savedName;
            string savedFolder;
            var title = string.Format(CodeExplorerUI.ResourceEditor_Title, displayName, projectName);
            using (var dialog = new NoteEditorDialog(title, NoteModule.ToMarkdown(opened), displayName, NoteModule.FolderOf(annotations), layouts.Get(layoutKey), preferences))
            {
                var result = dialog.ShowDialog();
                preferences.Save(_fileSystem);
                //How the window looked is remembered whether or not the note was saved.
                layouts.Set(layoutKey, dialog.Layout);
                layouts.Save(_fileSystem);
                if (result != DialogResult.OK || !dialog.IsModified)
                {
                    Logger.Debug("Note '{0}' closed without saving.", displayName);
                    return true;
                }

                saved = dialog.EditedText;
                savedName = dialog.EditedName;
                savedFolder = dialog.EditedFolder;
            }

            using (var component = _projects.Component(module))
            using (var codeModule = component.CodeModule)
            {
                //Someone else (an agent, say) changed the module while the window was open.
                if (codeModule.Content() != opened)
                {
                    Logger.Info("Note '{0}' changed while its editor was open; asking before overwriting.", displayName);
                    if (!_messageBox.ConfirmYesNo(string.Format(CodeExplorerUI.ResourceEditor_Stale, displayName), displayName, false))
                    {
                        Logger.Info("Note '{0}': the user kept the other change; nothing written.", displayName);
                        return true;
                    }
                }

                codeModule.Clear();
                codeModule.InsertLines(1, NoteModule.ToModuleCode(saved, NoteModule.WithFolder(NoteModule.WithMarker(annotations, savedName), savedFolder)));
                Logger.Info("Note '{0}' saved from the editor as '{1}' in folder '{2}' ({3} characters of markdown).", displayName, savedName, savedFolder, saved.Length);
            }

            return true;
        }

        private string ProjectFile(QualifiedModuleName module)
        {
            try
            {
                using (var project = _projects.Project(module.ProjectId))
                {
                    return project?.FileName;
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                //A project that was never saved has no file name.
                return null;
            }
        }
    }
}
