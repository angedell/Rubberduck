using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Rubberduck.Ship;
using Rubberduck.SmartIndenter;

namespace Rubberduck.Settings
{
    /// <summary>
    /// Creates and changes <c>rubberduck.json</c>. Changes are made to the text itself, so the comments and layout people gave the file stay as they are.
    /// </summary>
    public class ProjectConfigWriter
    {
        private readonly IFileSystem _fileSystem;

        public ProjectConfigWriter(IFileSystem fileSystem)
        {
            _fileSystem = fileSystem;
        }

        /// <summary>
        /// Writes a <c>rubberduck.json</c> with every setting and a comment for each, unless the folder already has one.
        /// </summary>
        /// <param name="indenter">The settings to write into the indenter section, normally the user's own.</param>
        /// <returns>The path of the file, and whether it was created now.</returns>
        public (string Path, bool Created) CreateDefault(string folder, IIndenterSettings indenter)
        {
            var path = _fileSystem.Path.Combine(folder, ProjectConfig.FileName);
            if (_fileSystem.File.Exists(path))
            {
                return (path, false);
            }

            _fileSystem.Directory.CreateDirectory(folder);
            _fileSystem.File.WriteAllText(path, ProjectConfigTemplate.Create(indenter));
            return (path, true);
        }

        /// <summary>
        /// Adds patterns to <c>sync.ignore</c>. Creates the file in <paramref name="defaultFolder"/> if <paramref name="configPath"/> is null.
        /// </summary>
        /// <returns>The path of the file that was written.</returns>
        public string AddIgnored(string configPath, string defaultFolder, IEnumerable<string> patterns)
        {
            var path = configPath ?? _fileSystem.Path.Combine(defaultFolder, ProjectConfig.FileName);
            var wanted = patterns.ToList();

            if (!_fileSystem.File.Exists(path))
            {
                _fileSystem.File.WriteAllText(path, ProjectConfigTemplate.Create(new IndenterSettings(true)));
            }

            var text = _fileSystem.File.ReadAllText(path);
            if (!TryAddToIgnoreList(text, wanted, out var updated))
            {
                updated = AddThroughTree(text, wanted);
            }

            _fileSystem.File.WriteAllText(path, updated);
            return path;
        }

        #region Editing the text

        //Edits the existing "sync": { "ignore": [ ... ] } array in place. Returns false if there is none.
        public static bool TryAddToIgnoreList(string text, IReadOnlyCollection<string> patterns, out string updated)
        {
            updated = text;
            var starts = LineStarts(text);
            int Index(JsonTextReader reader) => starts[reader.LineNumber - 1] + reader.LinePosition - 1;

            var open = -1;
            var close = -1;
            var lastElement = -1;
            var existing = new List<string>();
            try
            {
                using (var reader = new JsonTextReader(new System.IO.StringReader(text)))
                {
                    while (reader.Read())
                    {
                        if (!string.Equals(reader.Path, "sync.ignore", StringComparison.OrdinalIgnoreCase) && !reader.Path.StartsWith("sync.ignore[", StringComparison.OrdinalIgnoreCase))
                        {
                            if (open >= 0 && close < 0)
                            {
                                return false;
                            }

                            continue;
                        }

                        if (reader.TokenType == JsonToken.PropertyName)
                        {
                            continue;
                        }

                        if (reader.TokenType == JsonToken.StartArray && open < 0)
                        {
                            open = Index(reader);
                        }
                        else if (reader.TokenType == JsonToken.String)
                        {
                            existing.Add((string)reader.Value);
                            lastElement = Index(reader);
                        }
                        else if (reader.TokenType == JsonToken.EndArray)
                        {
                            close = Index(reader);
                            break;
                        }
                        else if (open < 0 || reader.TokenType != JsonToken.Comment)
                        {
                            return false;
                        }
                    }
                }
            }
            catch (JsonReaderException)
            {
                return false;
            }

            if (open < 0 || close < 0 || text[open] != '[' || text[close] != ']')
            {
                return false;
            }

            var added = patterns.Where(p => !existing.Contains(p)).Select(JsonConvert.ToString).ToList();
            if (added.Count == 0)
            {
                return true;
            }

            var newline = text.Contains("\r\n") ? "\r\n" : "\n";
            var lineStart = text.LastIndexOf('\n', Math.Max(open - 1, 0)) + 1;
            var arrayIndent = new string(text.Skip(lineStart).TakeWhile(c => c == ' ' || c == '\t').ToArray());
            var itemIndent = arrayIndent + "  ";

            if (lastElement >= 0)
            {
                //An array written over several lines gets its new entries on lines of their own, indented like the last one.
                var elementLineStart = text.LastIndexOf('\n', lastElement) + 1;
                var multiline = open < elementLineStart;
                var indent = new string(text.Skip(elementLineStart).TakeWhile(c => c == ' ' || c == '\t').ToArray());

                var insert = multiline
                    ? string.Concat(added.Select(item => "," + newline + indent + item))
                    : string.Concat(added.Select(item => ", " + item));
                updated = text.Insert(lastElement + 1, insert);
                return true;
            }

            var between = text.Substring(open + 1, close - open - 1);
            if (between.Trim().Length == 0)
            {
                var body = string.Join("," + newline, added.Select(item => itemIndent + item));
                updated = text.Substring(0, open + 1) + newline + body + newline + arrayIndent + text.Substring(close);
            }
            else
            {
                //Only comments inside; the entries go in front of them.
                var body = string.Join("," + newline, added.Select(item => itemIndent + item));
                updated = text.Insert(open + 1, newline + body);
            }

            return true;
        }

        /// <summary>
        /// Sets the <c>ship</c> section of the config text and leaves everything else as it is, comments and layout included. The section is added after the last one if the text has none.
        /// Throws <see cref="JsonReaderException"/> if the text is not a JSON object.
        /// </summary>
        public static string SetShip(string text, ShipSettings ship)
        {
            var newline = text.Contains("\r\n") ? "\r\n" : "\n";
            var starts = LineStarts(text);
            int Index(JsonTextReader reader) => starts[reader.LineNumber - 1] + reader.LinePosition - 1;

            int rootOpen = -1, rootClose = -1, shipOpen = -1, shipClose = -1, lastEnd = -1;
            using (var reader = new JsonTextReader(new System.IO.StringReader(text)))
            {
                while (reader.Read())
                {
                    if (reader.TokenType == JsonToken.Comment)
                    {
                        continue;
                    }

                    if (rootOpen < 0)
                    {
                        if (reader.TokenType != JsonToken.StartObject)
                        {
                            throw new JsonReaderException("The config must be a JSON object.", reader.Path, reader.LineNumber, reader.LinePosition, null);
                        }

                        rootOpen = lastEnd = Index(reader);
                    }
                    else if (reader.Depth == 0 && reader.TokenType == JsonToken.EndObject)
                    {
                        rootClose = Index(reader);
                        break;
                    }
                    else if (reader.Depth == 1 && reader.TokenType == JsonToken.PropertyName
                             && string.Equals((string)reader.Value, ShipSettingsJson.Key, StringComparison.OrdinalIgnoreCase))
                    {
                        do
                        {
                            reader.Read();
                        }
                        while (reader.TokenType == JsonToken.Comment);

                        if (reader.TokenType != JsonToken.StartObject)
                        {
                            throw new JsonReaderException("'ship' must be an object.", reader.Path, reader.LineNumber, reader.LinePosition, null);
                        }

                        shipOpen = Index(reader);
                        reader.Skip();
                        shipClose = lastEnd = Index(reader);
                    }
                    else
                    {
                        lastEnd = Index(reader);
                    }
                }
            }

            if (rootClose < 0)
            {
                throw new JsonReaderException("The config is not complete JSON.");
            }

            string IndentOf(int index) => new string(text.Skip(text.LastIndexOf('\n', Math.Max(index - 1, 0)) + 1).TakeWhile(c => c == ' ' || c == '\t').ToArray());

            if (shipOpen >= 0)
            {
                return text.Substring(0, shipOpen) + ShipSettingsJson.Section(ship, IndentOf(shipOpen), newline) + text.Substring(shipClose + 1);
            }

            var added = newline
                        + string.Concat(ShipSettingsJson.SectionComment.Select(line => "  // " + line + newline))
                        + "  \"" + ShipSettingsJson.Key + "\": " + ShipSettingsJson.Section(ship, "  ", newline);
            return lastEnd == rootOpen
                ? text.Substring(0, rootOpen + 1) + newline + added.Substring(newline.Length) + newline + text.Substring(rootOpen + 1)
                : text.Substring(0, lastEnd + 1) + "," + newline + added + text.Substring(lastEnd + 1);
        }

        private static List<int> LineStarts(string text)
        {
            var starts = new List<int> { 0 };
            for (var i = 0; i < text.Length; i++)
            {
                if (text[i] == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                {
                    i++;
                    starts.Add(i + 1);
                }
                else if (text[i] == '\r' || text[i] == '\n')
                {
                    starts.Add(i + 1);
                }
            }

            return starts;
        }

        //For a file that has no sync.ignore yet: the comments survive, but as /* */ comments.
        private static string AddThroughTree(string text, IEnumerable<string> patterns)
        {
            var root = JObject.Parse(text, new JsonLoadSettings { CommentHandling = CommentHandling.Load });

            var sync = root.Properties().FirstOrDefault(p => p.Name.Equals("sync", StringComparison.OrdinalIgnoreCase))?.Value as JObject;
            if (sync == null)
            {
                sync = new JObject();
                root["sync"] = sync;
            }

            var ignore = sync.Properties().FirstOrDefault(p => p.Name.Equals("ignore", StringComparison.OrdinalIgnoreCase))?.Value as JArray;
            if (ignore == null)
            {
                ignore = new JArray();
                sync["ignore"] = ignore;
            }

            foreach (var pattern in patterns)
            {
                if (!ignore.Any(item => item.Type == JTokenType.String && item.Value<string>() == pattern))
                {
                    ignore.Add(pattern);
                }
            }

            return root.ToString(Formatting.Indented) + Environment.NewLine;
        }

        #endregion
    }
}
