using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
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
