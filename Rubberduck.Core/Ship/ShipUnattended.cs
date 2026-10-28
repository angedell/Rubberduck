using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using Newtonsoft.Json;
using Rubberduck.Resources;

namespace Rubberduck.Ship
{
    /// <summary>
    /// The user's own switch for unattended ship, per document, in a small file of their own (next to <c>note-editor.json</c>), never in the project's rubberduck.json:
    /// an agent can edit that, and a switch that waives the agent's approval must not be in reach of the agent.
    /// </summary>
    public class ShipUnattended
    {
        public static readonly string FilePath = System.IO.Path.Combine(ApplicationConstants.RUBBERDUCK_FOLDER_PATH, "ship-unattended.json");

        //Windows paths ignore case.
        private Dictionary<string, bool> _documents = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

        [JsonProperty("documents")]
        public Dictionary<string, bool> Documents
        {
            get => _documents;
            set => _documents = new Dictionary<string, bool>(value ?? new Dictionary<string, bool>(), StringComparer.OrdinalIgnoreCase);
        }

        public bool IsOn(string documentPath) => documentPath != null && _documents.TryGetValue(documentPath, out var on) && on;

        public void Set(string documentPath, bool on)
        {
            if (documentPath == null)
            {
                return;
            }

            if (on)
            {
                _documents[documentPath] = true;
            }
            else
            {
                _documents.Remove(documentPath);
            }
        }

        /// <summary>The saved switches; none when there is no file or it cannot be read.</summary>
        public static ShipUnattended Load(IFileSystem fileSystem, string path = null)
        {
            try
            {
                path = path ?? FilePath;
                if (fileSystem.File.Exists(path))
                {
                    var loaded = JsonConvert.DeserializeObject<ShipUnattended>(fileSystem.File.ReadAllText(path));
                    if (loaded != null)
                    {
                        return loaded;
                    }
                }
            }
            catch (Exception exception) when (exception is JsonException || exception is System.IO.IOException || exception is UnauthorizedAccessException)
            {
                //No switch is the safe state.
            }

            return new ShipUnattended();
        }

        /// <exception cref="System.IO.IOException">The switch could not be kept; the caller should say so, since the user asked for it.</exception>
        public void Save(IFileSystem fileSystem, string path = null)
        {
            path = path ?? FilePath;
            fileSystem.Directory.CreateDirectory(fileSystem.Path.GetDirectoryName(path));
            fileSystem.File.WriteAllText(path, JsonConvert.SerializeObject(this, Formatting.Indented));
        }
    }
}
