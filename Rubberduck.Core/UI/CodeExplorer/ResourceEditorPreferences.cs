using System;
using System.IO.Abstractions;
using Newtonsoft.Json;
using Rubberduck.Resources;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>The font of the resource editor window. Kept in its own small file next to the settings, so the settings page need not know about it.</summary>
    public class ResourceEditorPreferences
    {
        public const string DefaultFontName = "Consolas";
        public const float DefaultFontSize = 10f;
        public const float MinFontSize = 6f;
        public const float MaxFontSize = 48f;

        public static readonly string FilePath = System.IO.Path.Combine(ApplicationConstants.RUBBERDUCK_FOLDER_PATH, "resource-editor.json");

        public string FontName { get; set; } = DefaultFontName;

        /// <summary>In points.</summary>
        public float FontSize { get; set; } = DefaultFontSize;

        public static float ClampSize(float size) => float.IsNaN(size) ? DefaultFontSize : Math.Max(MinFontSize, Math.Min(MaxFontSize, size));

        /// <summary>The saved preferences; the defaults when there are none or the file cannot be read.</summary>
        public static ResourceEditorPreferences Load(IFileSystem fileSystem, string path = null)
        {
            try
            {
                path = path ?? FilePath;
                if (fileSystem.File.Exists(path))
                {
                    var loaded = JsonConvert.DeserializeObject<ResourceEditorPreferences>(fileSystem.File.ReadAllText(path));
                    if (loaded != null)
                    {
                        loaded.FontName = string.IsNullOrWhiteSpace(loaded.FontName) ? DefaultFontName : loaded.FontName;
                        loaded.FontSize = ClampSize(loaded.FontSize);
                        return loaded;
                    }
                }
            }
            catch (Exception exception) when (exception is JsonException || exception is System.IO.IOException || exception is UnauthorizedAccessException)
            {
                //A font is not worth a failure: fall back to the default.
            }

            return new ResourceEditorPreferences();
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
                //Not saved; the choice lasts until the window closes.
            }
        }
    }
}
