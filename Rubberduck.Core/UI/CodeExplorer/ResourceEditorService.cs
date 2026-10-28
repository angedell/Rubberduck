using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using NLog;
using Rubberduck.Interaction;
using Rubberduck.PackageResources;
using Rubberduck.Resources.CodeExplorer;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>
    /// What the Code Explorer's resource commands do, with the dialogs they need: edit ribbon XML, add and remove pictures.
    /// The changes are staged in the <see cref="IPackageResourceService"/>; this only asks the user and reports what went wrong.
    /// </summary>
    public class ResourceEditorService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IPackageResourceService _resources;
        private readonly IMessageBox _messageBox;
        private readonly IFileSystemBrowserFactory _dialogFactory;
        private readonly IFileSystem _fileSystem;

        public ResourceEditorService(IPackageResourceService resources, IMessageBox messageBox, IFileSystemBrowserFactory dialogFactory, IFileSystem fileSystem)
        {
            _resources = resources;
            _messageBox = messageBox;
            _dialogFactory = dialogFactory;
            _fileSystem = fileSystem;
        }

        /// <summary>Opens the ribbon XML or the config in the editor window and stages what the user saves.</summary>
        public void EditText(string documentPath, string name, string projectName)
        {
            PackageResourceContent content;
            try
            {
                content = _resources.Read(documentPath, name);
            }
            catch (PackageResourceException exception)
            {
                Warn(exception.Message);
                return;
            }

            Edit(documentPath, name, projectName, content.Text, content.Hash);
        }

        /// <summary>Starts a ribbon XML from a template, for a document that has none. Office 2010 and later read customUI14.xml.</summary>
        public void AddRibbonXml(string documentPath, string projectName)
        {
            Edit(documentPath, PackageResourceNames.Ribbon2010, projectName, RibbonXml.Template(PackageResourceNames.Ribbon2010), null);
        }

        private void Edit(string documentPath, string name, string projectName, string xml, string hash)
        {
            string saved;
            var isConfig = PackageResourceNames.IsConfigName(name);
            Func<string, TextProblem> check = isConfig ? (Func<string, TextProblem>)ConfigJson.Check : text => RibbonXml.Check(name, text);
            var preferences = ResourceEditorPreferences.Load(_fileSystem);
            Func<string, string> format = isConfig ? (Func<string, string>)ConfigJson.Format : RibbonXml.Format;
            Func<string, string> linearize = isConfig ? (Func<string, string>)ConfigJson.Linearize : RibbonXml.Linearize;
            var colors = isConfig ? ResourceSyntaxColors.Json : ResourceSyntaxColors.Xml;
            using (var dialog = new ResourceTextEditorDialog(string.Format(CodeExplorerUI.ResourceEditor_Title, name, projectName), xml, check, format, linearize, colors, preferences))
            {
                var result = dialog.ShowDialog();
                preferences.Save(_fileSystem);

                //A new ribbon is saved even unedited: it is the template the user asked for.
                if (result != DialogResult.OK || (hash != null && !dialog.IsModified))
                {
                    return;
                }

                saved = dialog.EditedText;
            }

            try
            {
                try
                {
                    Save(documentPath, name, saved, hash);
                }
                catch (PackageResourceException exception) when (exception.Error == PackageResourceError.Stale)
                {
                    //Someone else (an agent, say) changed it while the window was open.
                    if (!_messageBox.ConfirmYesNo(string.Format(CodeExplorerUI.ResourceEditor_Stale, name), CodeExplorerUI.Resources_Caption, false))
                    {
                        return;
                    }

                    Save(documentPath, name, saved, _resources.Read(documentPath, name).Hash);
                }
            }
            catch (PackageResourceException exception)
            {
                Warn(exception.Message);
            }
        }

        private void Save(string documentPath, string name, string text, string hash)
        {
            if (PackageResourceNames.IsConfigName(name))
            {
                _resources.SetConfig(documentPath, text, hash);
            }
            else
            {
                _resources.SetXml(documentPath, name, text, hash);
            }
        }

        /// <summary>Asks for picture files and stages them. A name that is already used asks before replacing.</summary>
        public void AddImages(string documentPath)
        {
            string[] files;
            using (var dialog = _dialogFactory.CreateOpenFileDialog())
            {
                dialog.Filter = CodeExplorerUI.Resources_FilterPictures + " (" + string.Join(", ", PackageResourceNames.ImageExtensions.Select(e => "*" + e)) + ")|"
                                + string.Join(";", PackageResourceNames.ImageExtensions.Select(e => "*" + e));
                dialog.Multiselect = true;
                dialog.CheckFileExists = true;
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                files = dialog.FileNames;
            }

            var failures = new List<string>();
            foreach (var file in files)
            {
                var name = ResourceNameFor(_fileSystem.Path.GetFileName(file));
                try
                {
                    var bytes = _fileSystem.File.ReadAllBytes(file);
                    var existing = _resources.List(documentPath).FirstOrDefault(r => r.Name.Equals(name, System.StringComparison.OrdinalIgnoreCase) && r.Change != PackageResourceChange.Removed);
                    string hash = null;
                    if (existing != null)
                    {
                        if (!_messageBox.ConfirmYesNo(string.Format(CodeExplorerUI.Resources_ReplaceConfirm, name), CodeExplorerUI.Resources_Caption, false))
                        {
                            continue;
                        }

                        hash = _resources.Read(documentPath, name).Hash;
                    }

                    _resources.SetImage(documentPath, name, bytes, hash);
                }
                catch (PackageResourceException exception)
                {
                    failures.Add(string.Format(CodeExplorerUI.Resources_AddFailed, name, exception.Message));
                }
                catch (System.IO.IOException exception)
                {
                    failures.Add(string.Format(CodeExplorerUI.Resources_AddFailed, name, exception.Message));
                }
            }

            if (failures.Count > 0)
            {
                Warn(string.Join("\r\n", failures));
            }
        }

        /// <summary>Stages the removal of pictures, after asking.</summary>
        public bool RemoveImages(string documentPath, IReadOnlyCollection<string> names, bool confirm = true)
        {
            if (names.Count == 0)
            {
                return false;
            }

            if (confirm && !_messageBox.ConfirmYesNo(string.Format(CodeExplorerUI.Resources_RemoveConfirm, names.Count), CodeExplorerUI.Resources_Caption, false))
            {
                return false;
            }

            try
            {
                _resources.RemoveImages(documentPath, names);
                return true;
            }
            catch (PackageResourceException exception)
            {
                Warn(exception.Message);
                return false;
            }
        }

        public void Discard(string documentPath)
        {
            if (_messageBox.ConfirmYesNo(CodeExplorerUI.Resources_DiscardConfirm, CodeExplorerUI.Resources_Caption, false))
            {
                _resources.Discard(documentPath);
            }
        }

        //The file name becomes the picture's id in the ribbon XML, which must start with a letter or underscore and have no spaces.
        public static string ResourceNameFor(string fileName)
        {
            var name = Regex.Replace(fileName ?? string.Empty, @"[^A-Za-z0-9_.\-]", "_");
            return name.Length > 0 && !char.IsLetter(name[0]) && name[0] != '_' ? "img_" + name : name;
        }

        private void Warn(string message)
        {
            Logger.Info(message);
            _messageBox.NotifyWarn(message, CodeExplorerUI.Resources_Caption);
        }
    }
}
