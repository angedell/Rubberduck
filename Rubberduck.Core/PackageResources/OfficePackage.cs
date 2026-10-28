using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using InvalidDataException = System.IO.InvalidDataException;
using MemoryStream = System.IO.MemoryStream;
using Path = System.IO.Path;
using Stream = System.IO.Stream;

namespace Rubberduck.PackageResources
{
    /// <summary>
    /// Reads and rewrites the ribbon part of an Office Open XML package (xlsm, docm, pptm...): <c>customUI/customUI.xml</c>, <c>customUI/customUI14.xml</c>
    /// and the pictures in <c>customUI/images</c>, together with the relationships and content types that have to agree with them.
    /// Works on streams, never on the file itself, so the caller decides how the result replaces the original.
    /// </summary>
    internal static class OfficePackage
    {
        private const string ContentTypesPart = "[Content_Types].xml";
        private const string RootRelationships = "_rels/.rels";
        private const string RibbonRelationshipsFolder = "customUI/_rels/";
        private const string ImageRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        private const string Ribbon2007Relationship = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility";
        private const string Ribbon2010Relationship = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility";
        private const string ConfigRelationship = "http://rubberduck-vba.com/relationships/config";

        private static readonly XNamespace ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
        private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/package/2006/relationships";

        /// <summary>The resources in the package as it is on disk.</summary>
        public static List<PackageResource> List(Stream source)
        {
            var resources = new List<PackageResource>();
            using (var zip = Open(source))
            {
                foreach (var entry in zip.Entries)
                {
                    var resource = Describe(entry.FullName);
                    if (resource == null)
                    {
                        continue;
                    }

                    resource.Size = entry.Length;
                    resources.Add(resource);
                }
            }

            return resources;
        }

        /// <summary>The bytes of a resource, or null if the package has none by that name.</summary>
        public static byte[] Read(Stream source, string name)
        {
            using (var zip = Open(source))
            {
                var entry = zip.Entries.FirstOrDefault(e => string.Equals(e.FullName, PathOf(name), StringComparison.OrdinalIgnoreCase));
                if (entry == null)
                {
                    return null;
                }

                using (var stream = entry.Open())
                using (var buffer = new MemoryStream())
                {
                    stream.CopyTo(buffer);
                    return buffer.ToArray();
                }
            }
        }

        /// <summary>
        /// Copies the package from <paramref name="source"/> to <paramref name="target"/> with the edits applied. An edit with null content removes the resource.
        /// Relationships for the pictures, the ribbon parts' own relationship and the content types are brought in line.
        /// </summary>
        public static void Patch(Stream source, Stream target, IReadOnlyDictionary<string, byte[]> edits)
        {
            var editedPaths = new HashSet<string>(edits.Keys.Select(PathOf), StringComparer.OrdinalIgnoreCase);

            using (var input = Open(source))
            using (var output = new ZipArchive(target, ZipArchiveMode.Create, true))
            {
                var existing = input.Entries.Select(e => e.FullName).ToList();
                var finalRibbon = Final(existing, edits, PackageResourceKind.RibbonXml);
                var finalImages = Final(existing, edits, PackageResourceKind.Image);
                var configPresent = Final(existing, edits, PackageResourceKind.Config).Count > 0;

                var contentTypes = Xml(input, ContentTypesPart)
                    ?? throw new PackageResourceException(PackageResourceError.Unavailable, "The file is not an Office document package ([Content_Types].xml is missing).");
                Write(output, ContentTypesPart, UpdateContentTypes(contentTypes, finalImages, configPresent));
                Write(output, RootRelationships, UpdateRootRelationships(Xml(input, RootRelationships), finalRibbon, configPresent));

                var rewritten = new HashSet<string>(new[] { ContentTypesPart, RootRelationships }, StringComparer.OrdinalIgnoreCase);
                foreach (var ribbon in finalRibbon)
                {
                    var relationshipsPath = RibbonRelationshipsFolder + ribbon + ".rels";
                    Write(output, relationshipsPath, UpdateRibbonRelationships(Xml(input, relationshipsPath), finalImages));
                    rewritten.Add(relationshipsPath);
                }

                foreach (var entry in input.Entries)
                {
                    var path = entry.FullName;
                    if (path.EndsWith("/", StringComparison.Ordinal)
                        || rewritten.Contains(path)
                        || editedPaths.Contains(path)
                        || path.StartsWith(RibbonRelationshipsFolder, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var copy = output.CreateEntry(path, CompressionLevel.Optimal);
                    copy.LastWriteTime = entry.LastWriteTime;
                    using (var from = entry.Open())
                    using (var to = copy.Open())
                    {
                        from.CopyTo(to);
                    }
                }

                foreach (var edit in edits.Where(e => e.Value != null))
                {
                    var entry = output.CreateEntry(PathOf(edit.Key), CompressionLevel.Optimal);
                    using (var to = entry.Open())
                    {
                        to.Write(edit.Value, 0, edit.Value.Length);
                    }
                }
            }
        }

        public static string PathOf(string name)
            => (PackageResourceNames.IsImageName(name) ? PackageResourceNames.ImagesFolder : PackageResourceNames.RibbonFolder) + name;

        private static PackageResource Describe(string path)
        {
            if (path.StartsWith(PackageResourceNames.ImagesFolder, StringComparison.OrdinalIgnoreCase))
            {
                var name = path.Substring(PackageResourceNames.ImagesFolder.Length);
                return name.IndexOf('/') < 0 && PackageResourceNames.IsImageName(name)
                    ? new PackageResource { Name = name, Kind = PackageResourceKind.Image }
                    : null;
            }

            if (path.StartsWith(PackageResourceNames.RibbonFolder, StringComparison.OrdinalIgnoreCase))
            {
                var name = path.Substring(PackageResourceNames.RibbonFolder.Length);
                if (PackageResourceNames.IsConfigName(name))
                {
                    return new PackageResource { Name = PackageResourceNames.ConfigName, Kind = PackageResourceKind.Config };
                }

                return PackageResourceNames.IsRibbonName(name)
                    ? new PackageResource { Name = PackageResourceNames.RibbonNames.First(n => n.Equals(name, StringComparison.OrdinalIgnoreCase)), Kind = PackageResourceKind.RibbonXml }
                    : null;
            }

            return null;
        }

        //The names of the resources of a kind once the edits are applied.
        private static List<string> Final(IEnumerable<string> existingPaths, IReadOnlyDictionary<string, byte[]> edits, PackageResourceKind kind)
        {
            var names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var resource in existingPaths.Select(Describe).Where(r => r != null && r.Kind == kind))
            {
                names[resource.Name] = resource.Name;
            }

            foreach (var edit in edits.Where(e => PackageResourceNames.KindOf(e.Key) == kind))
            {
                if (edit.Value == null)
                {
                    names.Remove(edit.Key);
                }
                else
                {
                    names[edit.Key] = edit.Key;
                }
            }

            return names.Values.ToList();
        }

        private static ZipArchive Open(Stream source)
        {
            try
            {
                return new ZipArchive(source, ZipArchiveMode.Read, true);
            }
            catch (InvalidDataException exception)
            {
                throw new PackageResourceException(PackageResourceError.Unavailable, "The file is not an Office document package.", exception);
            }
        }

        private static XDocument Xml(ZipArchive zip, string path)
        {
            var entry = zip.Entries.FirstOrDefault(e => string.Equals(e.FullName, path, StringComparison.OrdinalIgnoreCase));
            if (entry == null)
            {
                return null;
            }

            using (var stream = entry.Open())
            {
                return XDocument.Load(stream);
            }
        }

        private static void Write(ZipArchive zip, string path, XDocument document)
        {
            var entry = zip.CreateEntry(path, CompressionLevel.Optimal);
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false }))
            {
                document.Save(writer);
            }
        }

        private static XDocument UpdateContentTypes(XDocument document, IEnumerable<string> images, bool config)
        {
            var root = document.Root;
            var extensions = new HashSet<string>(
                root.Elements(ContentTypes + "Default").Select(d => (string)d.Attribute("Extension")),
                StringComparer.OrdinalIgnoreCase);

            void Ensure(string extension)
            {
                if (!extensions.Add(extension))
                {
                    return;
                }

                //Defaults go before the overrides, as the schema orders them.
                var added = new XElement(ContentTypes + "Default",
                    new XAttribute("Extension", extension),
                    new XAttribute("ContentType", PackageResourceNames.ContentTypeOf("." + extension)));
                var lastDefault = root.Elements(ContentTypes + "Default").LastOrDefault();
                if (lastDefault != null)
                {
                    lastDefault.AddAfterSelf(added);
                }
                else
                {
                    root.AddFirst(added);
                }
            }

            Ensure("xml");
            if (config)
            {
                Ensure("json");
            }

            foreach (var extension in images.Select(i => Path.GetExtension(i).TrimStart('.')).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                Ensure(extension);
            }

            return document;
        }

        private static XDocument UpdateRootRelationships(XDocument document, IEnumerable<string> ribbons, bool config)
        {
            document = document ?? new XDocument(new XElement(Relationships + "Relationships"));
            var root = document.Root;

            //A relationship to a part that is gone makes Office report a damaged file.
            if (!config)
            {
                root.Elements(Relationships + "Relationship").Where(r => (string)r.Attribute("Type") == ConfigRelationship).ToList().ForEach(r => r.Remove());
            }

            if (config && !root.Elements(Relationships + "Relationship").Any(r => (string)r.Attribute("Type") == ConfigRelationship))
            {
                root.Add(new XElement(Relationships + "Relationship",
                    new XAttribute("Type", ConfigRelationship),
                    new XAttribute("Target", PackageResourceNames.RibbonFolder + PackageResourceNames.ConfigName),
                    new XAttribute("Id", UniqueId(root, "rdConfig"))));
            }

            foreach (var ribbon in ribbons)
            {
                var type = ribbon.Equals(PackageResourceNames.Ribbon2007, StringComparison.OrdinalIgnoreCase) ? Ribbon2007Relationship : Ribbon2010Relationship;
                if (root.Elements(Relationships + "Relationship").Any(r => (string)r.Attribute("Type") == type))
                {
                    continue;
                }

                root.Add(new XElement(Relationships + "Relationship",
                    new XAttribute("Type", type),
                    new XAttribute("Target", PackageResourceNames.RibbonFolder + ribbon),
                    new XAttribute("Id", UniqueId(root, "rdRibbon" + Path.GetFileNameWithoutExtension(ribbon)))));
            }

            return document;
        }

        //The ribbon XML refers to a picture by the id of a relationship, so every picture needs one, and none may point at a picture that is gone.
        private static XDocument UpdateRibbonRelationships(XDocument document, IReadOnlyCollection<string> images)
        {
            document = document ?? new XDocument(new XElement(Relationships + "Relationships"));
            var root = document.Root;
            var present = new HashSet<string>(images, StringComparer.OrdinalIgnoreCase);

            foreach (var stale in root.Elements(Relationships + "Relationship")
                .Where(r => (string)r.Attribute("Type") == ImageRelationshipType
                            && !present.Contains(Path.GetFileName((string)r.Attribute("Target") ?? string.Empty)))
                .ToList())
            {
                stale.Remove();
            }

            var linked = new HashSet<string>(
                root.Elements(Relationships + "Relationship").Select(r => Path.GetFileName((string)r.Attribute("Target") ?? string.Empty)),
                StringComparer.OrdinalIgnoreCase);

            foreach (var image in images.Where(i => !linked.Contains(i)))
            {
                root.Add(new XElement(Relationships + "Relationship",
                    new XAttribute("Type", ImageRelationshipType),
                    new XAttribute("Target", "images/" + image),
                    new XAttribute("Id", UniqueId(root, Path.GetFileNameWithoutExtension(image)))));
            }

            return document;
        }

        private static string UniqueId(XElement relationships, string wanted)
        {
            var taken = new HashSet<string>(relationships.Elements(Relationships + "Relationship").Select(r => (string)r.Attribute("Id")), StringComparer.Ordinal);
            var id = wanted;
            for (var counter = 2; taken.Contains(id); counter++)
            {
                id = wanted + counter;
            }

            return id;
        }
    }
}
