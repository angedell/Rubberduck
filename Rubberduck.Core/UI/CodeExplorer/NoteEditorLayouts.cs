using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using Newtonsoft.Json;
using Rubberduck.Resources;

namespace Rubberduck.UI.CodeExplorer
{
    public enum NoteEditorView
    {
        Split,
        Source,
        Preview
    }

    /// <summary>How the window of one note looks: which side is shown, how the room is shared, how far the preview is zoomed.</summary>
    public class NoteEditorLayout
    {
        public const double MinZoom = 50;
        public const double MaxZoom = 300;
        public const double DefaultZoom = 100;
        public const double DefaultSourceShare = 0.5;

        public NoteEditorView View { get; set; } = NoteEditorView.Split;

        /// <summary>The part of the width the markdown takes when both sides are shown, between 0.1 and 0.9.</summary>
        public double SourceShare { get; set; } = DefaultSourceShare;

        /// <summary>In percent.</summary>
        public double Zoom { get; set; } = DefaultZoom;

        public static double ClampZoom(double zoom) => double.IsNaN(zoom) ? DefaultZoom : Math.Max(MinZoom, Math.Min(MaxZoom, zoom));

        public static double ClampShare(double share) => double.IsNaN(share) ? DefaultSourceShare : Math.Max(0.1, Math.Min(0.9, share));

        internal NoteEditorLayout Sanitized()
        {
            return new NoteEditorLayout
            {
                View = Enum.IsDefined(typeof(NoteEditorView), View) ? View : NoteEditorView.Split,
                SourceShare = ClampShare(SourceShare),
                Zoom = ClampZoom(Zoom)
            };
        }
    }

    /// <summary>
    /// The layout of every note window, kept per note in a small file of the user's own, not in the workbook: how one reads a note is no business of the next person to open the file.
    /// A note is told apart by the file of its project and its module name.
    /// </summary>
    public class NoteEditorLayouts
    {
        public static readonly string FilePath = System.IO.Path.Combine(ApplicationConstants.RUBBERDUCK_FOLDER_PATH, "note-editor.json");

        //Ordinal and case-insensitive: Windows paths and VBA names both ignore case.
        private Dictionary<string, NoteEditorLayout> _notes = new Dictionary<string, NoteEditorLayout>(StringComparer.OrdinalIgnoreCase);

        [JsonProperty("notes")]
        public Dictionary<string, NoteEditorLayout> Notes
        {
            get => _notes;
            set => _notes = new Dictionary<string, NoteEditorLayout>(value ?? new Dictionary<string, NoteEditorLayout>(), StringComparer.OrdinalIgnoreCase);
        }

        public static string KeyFor(string projectFile, string projectName, string moduleName)
        {
            //A project that was never saved has no file: its name has to do.
            var project = string.IsNullOrWhiteSpace(projectFile) ? "?" + projectName : projectFile;
            return project + "|" + moduleName;
        }

        public NoteEditorLayout Get(string key)
        {
            return key != null && _notes.TryGetValue(key, out var layout) && layout != null ? layout.Sanitized() : new NoteEditorLayout();
        }

        public void Set(string key, NoteEditorLayout layout)
        {
            if (key != null && layout != null)
            {
                _notes[key] = layout.Sanitized();
            }
        }

        /// <summary>The saved layouts; none when there is no file or it cannot be read.</summary>
        public static NoteEditorLayouts Load(IFileSystem fileSystem, string path = null)
        {
            try
            {
                path = path ?? FilePath;
                if (fileSystem.File.Exists(path))
                {
                    var loaded = JsonConvert.DeserializeObject<NoteEditorLayouts>(fileSystem.File.ReadAllText(path));
                    if (loaded != null)
                    {
                        return loaded;
                    }
                }
            }
            catch (Exception exception) when (exception is JsonException || exception is System.IO.IOException || exception is UnauthorizedAccessException)
            {
                //A window layout is not worth a failure: start from the defaults.
            }

            return new NoteEditorLayouts();
        }

        public void Save(IFileSystem fileSystem, string path = null)
        {
            try
            {
                path = path ?? FilePath;
                fileSystem.Directory.CreateDirectory(fileSystem.Path.GetDirectoryName(path));
                fileSystem.File.WriteAllText(path, JsonConvert.SerializeObject(this, Formatting.Indented));
            }
            catch (Exception exception) when (exception is System.IO.IOException || exception is UnauthorizedAccessException)
            {
                //Not saved; the layout lasts until the window closes.
            }
        }
    }
}
