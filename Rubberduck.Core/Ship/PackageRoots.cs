using System;
using System.Collections.Generic;
using Stream = System.IO.Stream;
using StreamReader = System.IO.StreamReader;
using StringReader = System.IO.StringReader;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Rubberduck.Ship
{
    /// <summary>A member name found in a part of the document's package, and where.</summary>
    public class PackageRootHit
    {
        public string Name { get; set; }
        /// <summary>The part, and for the ribbon the attribute: <c>customUI/customUI14.xml onAction</c>.</summary>
        public string Source { get; set; }
    }

    /// <summary>
    /// Roots that live in the document and not in the code (ADR 0004 item 3): the ribbon's callbacks (every word in the attribute values of <c>customUI/customUI*.xml</c>),
    /// the macros of shapes, controls and charts, defined names, the formulas of worksheets, tables, pivot tables and caches, and query tables. A member name that appears as a whole word keeps the member:
    /// deliberately over-cautious, because a name in a formula or a macro assignment cannot be resolved without Excel.
    /// Noise is accepted by design: sheet names, pivot item values and any other word in these parts that equals a member name keep that member.
    /// This is a pure function of the text of the parts; <see cref="Find(Stream, IEnumerable{string})"/> only reads them out of the zip.
    /// </summary>
    public static class PackageRoots
    {
        private static readonly Regex RibbonPart = new Regex(@"^customUI/customUI[^/]*\.xml$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex CodePart = new Regex(@"^(xl/(worksheets|chartsheets|drawings|charts|ctrlProps|macrosheets|dialogsheets)/[^/]+\.(xml|vml)|xl/(tables|pivotTables|queryTables|slicerCaches|slicers|externalLinks)/[^/]+\.xml|xl/connections\.xml|xl/pivotCache[^/]*/(?!pivotCacheRecords)[^/]+\.xml|xl/workbook\.xml)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex Word = new Regex(@"[\p{L}_][\p{L}\p{Nd}_]*", RegexOptions.CultureInvariant);

        //Cell values and the attributes that only locate a cell are data. A formula is code.
        private static readonly HashSet<string> SheetCodeElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "f", "formula", "formula1", "formula2", "definedName" };
        private static readonly HashSet<string> SheetDataElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "c", "row", "col", "dimension", "v", "t", "is", "si" };

        public static bool IsRelevantPart(string path) => RibbonPart.IsMatch(path) || CodePart.IsMatch(path);

        /// <summary>Reads the relevant parts of a zip package (the document's file) and searches them.</summary>
        public static List<PackageRootHit> Find(Stream package, IEnumerable<string> names)
        {
            var parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            using (var zip = new ZipArchive(package, ZipArchiveMode.Read, true))
            {
                foreach (var entry in zip.Entries.Where(e => IsRelevantPart(e.FullName)))
                {
                    using (var stream = entry.Open())
                    using (var reader = new StreamReader(stream, Encoding.UTF8, true))
                    {
                        parts[entry.FullName] = reader.ReadToEnd();
                    }
                }
            }

            return Find(parts, names);
        }

        /// <summary>Searches parts given as path and text. Names are matched without regard to case; the hits carry the name as it was given.</summary>
        public static List<PackageRootHit> Find(IReadOnlyDictionary<string, string> parts, IEnumerable<string> names)
        {
            var wanted = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var name in names.Where(n => !string.IsNullOrEmpty(n)))
            {
                wanted[name] = name;
            }

            var hits = new List<PackageRootHit>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            void Hit(string word, string source)
            {
                if (wanted.TryGetValue(word, out var name) && seen.Add(name + "|" + source))
                {
                    hits.Add(new PackageRootHit { Name = name, Source = source });
                }
            }

            foreach (var part in parts.OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase))
            {
                var isRibbon = RibbonPart.IsMatch(part.Key);
                if (!isRibbon && !CodePart.IsMatch(part.Key))
                {
                    continue;
                }

                var isSheet = part.Key.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase) || part.Key.StartsWith("xl/chartsheets/", StringComparison.OrdinalIgnoreCase);
                if (!TryReadXml(part.Key, part.Value, isRibbon, isSheet, Hit))
                {
                    foreach (Match word in Word.Matches(part.Value))
                    {
                        Hit(word.Value, part.Key);
                    }
                }
            }

            return hits;
        }

        //False when the text is not well-formed XML: the caller then searches it as plain text.
        private static bool TryReadXml(string path, string text, bool isRibbon, bool isSheet, Action<string, string> hit)
        {
            var found = new List<(string Word, string Source)>();
            try
            {
                var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, IgnoreWhitespace = true };
                using (var reader = XmlReader.Create(new StringReader(text), settings))
                {
                    var element = string.Empty;
                    while (reader.Read())
                    {
                        switch (reader.NodeType)
                        {
                            case XmlNodeType.Element:
                                element = reader.LocalName;
                                if (reader.HasAttributes && !(isSheet && SheetDataElements.Contains(element)))
                                {
                                    while (reader.MoveToNextAttribute())
                                    {
                                        if (reader.LocalName == "xmlns" || reader.Prefix == "xmlns")
                                        {
                                            continue;
                                        }

                                        var value = reader.Value.Trim();
                                        if (isRibbon)
                                        {
                                            //onAction="Module1.OnRun", "'Book.xlam'!OnRun", "Module1.OnRun 1": every word is a candidate.
                                            found.AddRange(Word.Matches(value).Cast<Match>().Select(m => (m.Value, path + " " + reader.LocalName)));
                                        }
                                        else
                                        {
                                            found.AddRange(Word.Matches(value).Cast<Match>().Select(m => (m.Value, path)));
                                        }
                                    }

                                    reader.MoveToElement();
                                }

                                break;
                            case XmlNodeType.Text:
                            case XmlNodeType.CDATA:
                                if (!isRibbon && (!isSheet || SheetCodeElements.Contains(element)))
                                {
                                    found.AddRange(Word.Matches(reader.Value).Cast<Match>().Select(m => (m.Value, path)));
                                }

                                break;
                        }
                    }
                }
            }
            catch (XmlException)
            {
                return false;
            }

            foreach (var (word, source) in found)
            {
                hit(word, source);
            }

            return true;
        }
    }
}
