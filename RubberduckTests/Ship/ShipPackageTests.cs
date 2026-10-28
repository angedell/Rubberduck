using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using NUnit.Framework;
using Rubberduck.PackageResources;

namespace RubberduckTests.Ship
{
    /// <summary>RemoveConfig takes the part out of the package with the existing Patch: the part, and the root relationship that points at it.</summary>
    [TestFixture]
    [Category("Ship")]
    public class ShipPackageTests
    {
        private const string ContentTypes = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"json\" ContentType=\"application/json\"/></Types>";

        private const string RootRelationships = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                                                 + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                                                 + "<Relationship Id=\"rdRibbon\" Type=\"http://schemas.microsoft.com/office/2007/relationships/ui/extensibility\" Target=\"customUI/customUI14.xml\"/>"
                                                 + "<Relationship Id=\"rdConfig\" Type=\"http://rubberduck-vba.com/relationships/config\" Target=\"customUI/rubberduck.json\"/>"
                                                 + "</Relationships>";

        private static MemoryStream Package()
        {
            var stream = new MemoryStream();
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, true))
            {
                void Add(string path, string content)
                {
                    using (var writer = new StreamWriter(zip.CreateEntry(path).Open()))
                    {
                        writer.Write(content);
                    }
                }

                Add("[Content_Types].xml", ContentTypes);
                Add("_rels/.rels", RootRelationships);
                Add("xl/workbook.xml", "<workbook/>");
                Add("customUI/customUI14.xml", "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"/>");
                Add("customUI/rubberduck.json", "{}");
            }

            stream.Position = 0;
            return stream;
        }

        [Test]
        public void RemovingTheConfigTakesOutThePartAndItsRelationshipAndNothingElse()
        {
            var output = new MemoryStream();
            OfficePackage.Patch(Package(), output, new Dictionary<string, byte[]> { { "rubberduck.json", null } });
            output.Position = 0;

            using (var zip = new ZipArchive(output, ZipArchiveMode.Read))
            {
                var names = zip.Entries.Select(e => e.FullName).ToList();
                CollectionAssert.DoesNotContain(names, "customUI/rubberduck.json");
                CollectionAssert.Contains(names, "customUI/customUI14.xml");
                CollectionAssert.Contains(names, "xl/workbook.xml");

                XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";
                using (var stream = zip.GetEntry("_rels/.rels").Open())
                {
                    var targets = XDocument.Load(stream).Root.Elements(ns + "Relationship").Select(r => (string)r.Attribute("Target")).ToList();
                    CollectionAssert.AreEquivalent(new[] { "xl/workbook.xml", "customUI/customUI14.xml" }, targets);
                }
            }
        }
    }
}
