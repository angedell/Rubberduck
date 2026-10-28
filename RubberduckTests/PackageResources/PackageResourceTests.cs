using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Abstractions;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using NUnit.Framework;
using Rubberduck.PackageResources;
using Rubberduck.UI.CodeExplorer;

namespace RubberduckTests.PackageResources
{
    /// <summary>Real files in a temp folder: the locking and replacing that matter here are the file system's.</summary>
    [TestFixture]
    [Category("PackageResources")]
    public class PackageResourceServiceTests
    {
        public const string Ribbon = "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><ribbon/></customUI>";
        public static readonly byte[] Png = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 1, 2, 3 };

        private string _folder;
        private string _document;
        private PackageResourceService _service;

        [SetUp]
        public void SetUp()
        {
            _folder = Path.Combine(Path.GetTempPath(), "rd-resources-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_folder);
            _document = Path.Combine(_folder, "Book.xlsm");
            WriteDocument(_document, withRibbon: true);
            _service = NewService();
        }

        [TearDown]
        public void TearDown()
        {
            try
            {
                Directory.Delete(_folder, true);
            }
            catch (IOException)
            {
            }
        }

        private PackageResourceService NewService()
            => new PackageResourceService(new FileSystem(), Path.Combine(_folder, "journal"), Path.Combine(_folder, "backup"));

        public static void WriteDocument(string path, bool withRibbon)
        {
            using (var stream = File.Create(path))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                void Add(string name, string text) => Add2(zip, name, Encoding.UTF8.GetBytes(text));

                Add("[Content_Types].xml",
                    "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                    "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.ms-excel.sheet.macroEnabled.main+xml\"/></Types>");
                Add("_rels/.rels",
                    "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                    (withRibbon ? "<Relationship Id=\"rId2\" Type=\"http://schemas.microsoft.com/office/2007/relationships/ui/extensibility\" Target=\"customUI/customUI14.xml\"/>" : string.Empty) +
                    "</Relationships>");
                Add("xl/workbook.xml", "<workbook/>");

                if (withRibbon)
                {
                    Add("customUI/customUI14.xml", Ribbon);
                    Add("customUI/_rels/customUI14.xml.rels",
                        "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"logo\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"images/logo.png\"/></Relationships>");
                    Add2(zip, "customUI/images/logo.png", Png);
                }
            }
        }

        private static void Add2(ZipArchive zip, string name, byte[] bytes)
        {
            using (var entry = zip.CreateEntry(name).Open())
            {
                entry.Write(bytes, 0, bytes.Length);
            }
        }

        private static string EntryText(string path, string name)
        {
            using (var zip = ZipFile.OpenRead(path))
            using (var reader = new StreamReader(zip.GetEntry(name).Open()))
            {
                return reader.ReadToEnd();
            }
        }

        private static bool HasEntry(string path, string name)
        {
            using (var zip = ZipFile.OpenRead(path))
            {
                return zip.GetEntry(name) != null;
            }
        }

        [Test]
        public void List_ShowsRibbonXmlAndPictures_RibbonFirst()
        {
            var listing = _service.List(_document);

            CollectionAssert.AreEqual(new[] { "customUI14.xml", "logo.png" }, listing.Select(r => r.Name).ToArray());
            Assert.AreEqual(PackageResourceKind.RibbonXml, listing[0].Kind);
            Assert.AreEqual(PackageResourceKind.Image, listing[1].Kind);
            Assert.IsTrue(listing.All(r => r.Change == PackageResourceChange.None));
        }

        [Test]
        public void IsSupported_OnlyForSavedOpenXmlDocuments()
        {
            Assert.IsTrue(_service.IsSupported(_document));
            Assert.IsFalse(_service.IsSupported(Path.Combine(_folder, "Old.xls")));
            Assert.IsFalse(_service.IsSupported(Path.Combine(_folder, "Missing.xlsm")));
            Assert.IsFalse(_service.IsSupported(null));
            Assert.IsFalse(_service.IsSupported(string.Empty));
        }

        [Test]
        public void SetXml_StagesTheChange_AndLeavesTheFileAlone()
        {
            var before = File.ReadAllBytes(_document);
            var hash = _service.Read(_document, "customUI14.xml").Hash;
            var edited = Ribbon.Replace("<ribbon/>", "<ribbon><tabs/></ribbon>");

            _service.SetXml(_document, "customUI14.xml", edited, hash);

            CollectionAssert.AreEqual(before, File.ReadAllBytes(_document));
            var read = _service.Read(_document, "customUI14.xml");
            Assert.AreEqual(edited, read.Text);
            Assert.AreEqual(PackageResourceChange.Modified, read.Change);
            Assert.IsTrue(_service.HasPending(_document));
        }

        [Test]
        public void StagedChanges_SurviveANewServiceInstance()
        {
            _service.SetXml(_document, "customUI14.xml", Ribbon.Replace("<ribbon/>", "<ribbon><tabs/></ribbon>"), _service.Read(_document, "customUI14.xml").Hash);

            var later = NewService();

            StringAssert.Contains("<tabs/>", later.Read(_document, "customUI14.xml").Text);
            CollectionAssert.Contains(later.PendingDocuments(), _document);
        }

        [Test]
        public void SetXml_RejectsWhatOfficeWouldIgnore()
        {
            var hash = _service.Read(_document, "customUI14.xml").Hash;

            Assert.AreEqual(PackageResourceError.BadRequest, Assert.Throws<PackageResourceException>(() => _service.SetXml(_document, "customUI14.xml", "<customUI", hash)).Error);
            Assert.AreEqual(PackageResourceError.BadRequest, Assert.Throws<PackageResourceException>(() => _service.SetXml(_document, "customUI14.xml", "<other/>", hash)).Error);
            Assert.AreEqual(PackageResourceError.BadRequest, Assert.Throws<PackageResourceException>(() =>
                _service.SetXml(_document, "customUI14.xml", "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\"/>", hash)).Error);
            Assert.AreEqual(PackageResourceError.BadRequest, Assert.Throws<PackageResourceException>(() => _service.SetXml(_document, "notes.xml", Ribbon, null)).Error);
            Assert.IsFalse(_service.HasPending(_document));
        }

        [Test]
        public void SetXml_NeedsTheHashOfWhatItReplaces_AndChecksIt()
        {
            Assert.AreEqual(PackageResourceError.HashRequired, Assert.Throws<PackageResourceException>(() => _service.SetXml(_document, "customUI14.xml", Ribbon, null)).Error);
            Assert.AreEqual(PackageResourceError.Stale, Assert.Throws<PackageResourceException>(() => _service.SetXml(_document, "customUI14.xml", Ribbon, "0000")).Error);
        }

        [Test]
        public void SetXml_RewritesTheDeclaredEncodingToUtf8()
        {
            var hash = _service.Read(_document, "customUI14.xml").Hash;

            _service.SetXml(_document, "customUI14.xml", "<?xml version=\"1.0\" encoding=\"UTF-16\"?>" + Ribbon, hash);

            StringAssert.Contains("encoding=\"UTF-8\"", _service.Read(_document, "customUI14.xml").Text);
        }

        [Test]
        public void SetImage_AddsAPicture_AndChecksItIsOne()
        {
            _service.SetImage(_document, "new.png", Png, null);

            var added = _service.List(_document).Single(r => r.Name == "new.png");
            Assert.AreEqual(PackageResourceChange.Added, added.Change);
            Assert.AreEqual(PackageResourceError.BadRequest, Assert.Throws<PackageResourceException>(() => _service.SetImage(_document, "fake.png", new byte[] { 1, 2, 3 }, null)).Error);
            Assert.AreEqual(PackageResourceError.BadRequest, Assert.Throws<PackageResourceException>(() => _service.SetImage(_document, "1bad name.png", Png, null)).Error);
            Assert.AreEqual(PackageResourceError.BadRequest, Assert.Throws<PackageResourceException>(() => _service.SetImage(_document, "x.exe", Png, null)).Error);
        }

        [Test]
        public void SetImage_ReplacingNeedsTheHash()
        {
            Assert.AreEqual(PackageResourceError.HashRequired, Assert.Throws<PackageResourceException>(() => _service.SetImage(_document, "logo.png", Png, null)).Error);

            _service.SetImage(_document, "LOGO.png", Png.Concat(new byte[] { 9 }).ToArray(), _service.Read(_document, "logo.png").Hash);

            //The picture that is in the file keeps its name; no second one appears.
            Assert.AreEqual(1, _service.List(_document).Count(r => r.Kind == PackageResourceKind.Image));
            Assert.AreEqual(PackageResourceChange.Modified, _service.List(_document).Single(r => r.Kind == PackageResourceKind.Image).Change);
        }

        [Test]
        public void SetImage_NeedsRibbonXmlToBeThere()
        {
            var bare = Path.Combine(_folder, "Bare.xlsm");
            WriteDocument(bare, withRibbon: false);

            Assert.AreEqual(PackageResourceError.Conflict, Assert.Throws<PackageResourceException>(() => _service.SetImage(bare, "a.png", Png, null)).Error);

            _service.SetXml(bare, "customUI14.xml", Ribbon, null);
            Assert.DoesNotThrow(() => _service.SetImage(bare, "a.png", Png, null));
        }

        [Test]
        public void RemoveImages_StagesTheRemoval_AndDropsWhatWasOnlyStaged()
        {
            _service.SetImage(_document, "new.png", Png, null);

            _service.RemoveImages(_document, new[] { "logo.png", "new.png" });

            var listing = _service.List(_document);
            Assert.AreEqual(PackageResourceChange.Removed, listing.Single(r => r.Name == "logo.png").Change);
            Assert.IsFalse(listing.Any(r => r.Name == "new.png"));
        }

        [Test]
        public void RemoveImages_RefusesRibbonXml_AndUnknownNames_AndChangesNothing()
        {
            Assert.AreEqual(PackageResourceError.BadRequest, Assert.Throws<PackageResourceException>(() => _service.RemoveImages(_document, new[] { "customUI14.xml" })).Error);
            Assert.AreEqual(PackageResourceError.NotFound, Assert.Throws<PackageResourceException>(() => _service.RemoveImages(_document, new[] { "logo.png", "nope.png" })).Error);
            Assert.IsFalse(_service.HasPending(_document));
        }

        [Test]
        public void Discard_ThrowsStagedChangesAway()
        {
            _service.SetImage(_document, "new.png", Png, null);

            _service.Discard(_document);

            Assert.IsFalse(_service.HasPending(_document));
            Assert.IsFalse(_service.List(_document).Any(r => r.Name == "new.png"));
        }

        [Test]
        public void Changed_IsRaisedForStagingDiscardingAndApplying()
        {
            var raised = 0;
            _service.Changed += (sender, args) => raised++;

            _service.SetImage(_document, "new.png", Png, null);
            _service.Discard(_document);
            _service.SetImage(_document, "new.png", Png, null);
            _service.TryApply(_document, out _);

            Assert.AreEqual(4, raised);
        }

        [Test]
        public void TryApply_WritesTheChangesIntoTheFile_AndKeepsItConsistent()
        {
            var edited = Ribbon.Replace("<ribbon/>", "<ribbon><tabs/></ribbon>");
            _service.SetXml(_document, "customUI14.xml", edited, _service.Read(_document, "customUI14.xml").Hash);
            _service.SetImage(_document, "icon.jpg", new byte[] { 0xFF, 0xD8, 0xFF, 1 }, null);
            _service.RemoveImages(_document, new[] { "logo.png" });

            var outcome = _service.TryApply(_document, out _);

            Assert.AreEqual(PackageApplyOutcome.Applied, outcome);
            Assert.AreEqual(edited, EntryText(_document, "customUI/customUI14.xml"));
            Assert.IsTrue(HasEntry(_document, "customUI/images/icon.jpg"));
            Assert.IsFalse(HasEntry(_document, "customUI/images/logo.png"));
            Assert.IsTrue(HasEntry(_document, "xl/workbook.xml"), "the rest of the package is carried over");

            var relationships = EntryText(_document, "customUI/_rels/customUI14.xml.rels");
            StringAssert.Contains("images/icon.jpg", relationships);
            StringAssert.Contains("Id=\"icon\"", relationships);
            StringAssert.DoesNotContain("logo.png", relationships);
            StringAssert.Contains("Extension=\"jpg\"", EntryText(_document, "[Content_Types].xml"));
            StringAssert.Contains("customUI/customUI14.xml", EntryText(_document, "_rels/.rels"));

            Assert.IsFalse(_service.HasPending(_document));
            Assert.IsFalse(File.Exists(_document + ".rd-resources.tmp"));
            Assert.AreEqual(1, Directory.GetFiles(Path.Combine(_folder, "backup"), "Book.xlsm", SearchOption.AllDirectories).Length, "the original is kept");
        }

        [Test]
        public void TryApply_AddsTheRibbonRelationship_WhenTheDocumentHadNone()
        {
            var bare = Path.Combine(_folder, "Bare.xlsm");
            WriteDocument(bare, withRibbon: false);
            _service.SetXml(bare, "customUI14.xml", Ribbon, null);
            _service.SetImage(bare, "a.png", Png, null);

            Assert.AreEqual(PackageApplyOutcome.Applied, _service.TryApply(bare, out _));

            StringAssert.Contains("ui/extensibility", EntryText(bare, "_rels/.rels"));
            StringAssert.Contains("images/a.png", EntryText(bare, "customUI/_rels/customUI14.xml.rels"));
            StringAssert.Contains("Extension=\"png\"", EntryText(bare, "[Content_Types].xml"));
            Assert.AreEqual(new[] { "customUI14.xml", "a.png" }.OrderBy(n => n), _service.List(bare).Select(r => r.Name).OrderBy(n => n));
        }

        [Test]
        public void TryApply_WhileTheDocumentIsOpen_ChangesNothing_AndKeepsTheStagedChanges()
        {
            _service.SetImage(_document, "new.png", Png, null);
            var before = File.ReadAllBytes(_document);

            //What the application does: has the file open, and allows others to read it.
            using (new FileStream(_document, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                Assert.AreEqual(PackageApplyOutcome.Busy, _service.TryApply(_document, out var message));
                StringAssert.Contains("open", message);

                //Reading, as the tree and the API do, still works.
                Assert.AreEqual(2, _service.List(_document).Count(r => r.Name == "logo.png" || r.Name == "new.png"));
            }

            CollectionAssert.AreEqual(before, File.ReadAllBytes(_document));
            Assert.IsTrue(_service.HasPending(_document));
            Assert.IsFalse(File.Exists(_document + ".rd-resources.tmp"));

            Assert.AreEqual(PackageApplyOutcome.Applied, _service.TryApply(_document, out _), "applies once the document is closed");
        }

        [Test]
        public void TryApply_WithNothingStaged_DoesNothing()
        {
            Assert.AreEqual(PackageApplyOutcome.NothingPending, _service.TryApply(_document, out _));
        }

        [Test]
        public void TryApply_OfADocumentThatIsGone_FailsAndKeepsTheChanges()
        {
            _service.SetImage(_document, "new.png", Png, null);
            File.Delete(_document);

            Assert.AreEqual(PackageApplyOutcome.Failed, _service.TryApply(_document, out var message));
            StringAssert.Contains("no longer exists", message);
            Assert.IsTrue(_service.HasPending(_document));
        }

        [Test]
        public void Applier_AppliesWhatIsFree_AndLeavesWhatIsOpen()
        {
            var open = Path.Combine(_folder, "Open.xlsm");
            WriteDocument(open, withRibbon: true);
            _service.SetImage(_document, "new.png", Png, null);
            _service.SetImage(open, "new.png", Png, null);

            using (new FileStream(open, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                new PackageResourceApplier(_service, new Moq.Mock<Rubberduck.VBEditor.Events.IVbeEvents>().Object)
                    .ApplyPending(TimeSpan.Zero, System.Threading.CancellationToken.None);

                Assert.IsFalse(_service.HasPending(_document));
                Assert.IsTrue(_service.HasPending(open));
            }
        }
    }

    [TestFixture]
    [Category("PackageResources")]
    public class RibbonXmlTests
    {
        [Test]
        public void Check_AcceptsTheRightRootInTheRightNamespace()
        {
            Assert.IsNull(RibbonXml.Check("customUI14.xml", PackageResourceServiceTests.Ribbon));
            Assert.IsNull(RibbonXml.Check("customUI.xml", "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\"/>"));
        }

        [Test]
        public void Check_SaysWhereTheXmlIsBroken()
        {
            var problem = RibbonXml.Check("customUI14.xml", "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">\r\n  <ribbon>\r\n</customUI>");

            Assert.IsNotNull(problem);
            Assert.AreEqual(3, problem.Line);
        }

        [Test]
        public void Check_RefusesTheWrongRootOrNamespace()
        {
            Assert.IsNotNull(RibbonXml.Check("customUI14.xml", "<root/>"));
            Assert.IsNotNull(RibbonXml.Check("customUI14.xml", "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\"/>"));
            Assert.IsNotNull(RibbonXml.Check("customUI14.xml", " "));
        }

        [Test]
        public void Format_Indents_AndKeepsTheDeclarationAndComments()
        {
            var formatted = RibbonXml.Format("<?xml version=\"1.0\" encoding=\"UTF-8\"?><customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><!-- hi --><ribbon><tabs/></ribbon></customUI>");

            StringAssert.StartsWith("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n<customUI", formatted);
            StringAssert.Contains("\r\n  <!-- hi -->", formatted);
            StringAssert.Contains("\r\n    <tabs />", formatted);
        }

        [Test]
        public void Template_IsValidForBothNames()
        {
            Assert.IsNull(RibbonXml.Check("customUI14.xml", RibbonXml.Template("customUI14.xml")));
            Assert.IsNull(RibbonXml.Check("customUI.xml", RibbonXml.Template("customUI.xml")));
        }

        [Test]
        public void Linearize_PutsItOnOneLine_AndKeepsTheDeclarationAndComments()
        {
            var linear = RibbonXml.Linearize("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">\r\n  <!-- hi -->\r\n  <ribbon>\r\n    <tabs />\r\n  </ribbon>\r\n</customUI>");

            Assert.AreEqual("<?xml version=\"1.0\" encoding=\"UTF-8\"?><customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><!-- hi --><ribbon><tabs /></ribbon></customUI>", linear);
        }

        [Test]
        public void Linearize_ThenFormat_GivesTheFormattedTextBack()
        {
            var formatted = RibbonXml.Format(RibbonXml.Template("customUI14.xml"));

            Assert.AreEqual(formatted, RibbonXml.Format(RibbonXml.Linearize(formatted)));
        }

        [Test]
        public void Linearize_OfBrokenXml_Throws()
            => Assert.Throws<XmlException>(() => RibbonXml.Linearize("<customUI><ribbon></customUI>"));

        [Test]
        public void Check_RefusesTwoControlsWithTheSameId_AndSaysWhere()
        {
            var problem = RibbonXml.Check("customUI14.xml",
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">\r\n  <ribbon>\r\n    <tabs>\r\n      <tab id=\"a\"><group id=\"g\"><button id=\"a\" /></group></tab>\r\n    </tabs>\r\n  </ribbon>\r\n</customUI>");

            Assert.IsNotNull(problem);
            Assert.IsTrue(problem.Blocking);
            Assert.AreEqual(4, problem.Line);
            StringAssert.Contains("\"a\"", problem.Message);
        }
    }

    [TestFixture]
    [Category("PackageResources")]
    public class ConfigJsonLayoutTests
    {
        [Test]
        public void Format_IndentsByTwo_KeepingTheOrder()
        {
            var formatted = ConfigJson.Format("{\"b\":1,\"a\":{\"c\":[1,2]}}");

            Assert.AreEqual("{\r\n  \"b\": 1,\r\n  \"a\": {\r\n    \"c\": [\r\n      1,\r\n      2\r\n    ]\r\n  }\r\n}", formatted);
        }

        [Test]
        public void Linearize_PutsItOnOneLine()
            => Assert.AreEqual("{\"b\":1,\"a\":{\"c\":[1,2]}}", ConfigJson.Linearize("{\r\n  \"b\": 1,\r\n  \"a\": { \"c\": [1, 2] }\r\n}"));

        [Test]
        public void Format_AndLinearize_RefuseToDropComments()
        {
            const string commented = "{\r\n  // note\r\n  \"a\": 1\r\n}";

            StringAssert.Contains("comments", Assert.Throws<Newtonsoft.Json.JsonReaderException>(() => ConfigJson.Format(commented)).Message);
            Assert.Throws<Newtonsoft.Json.JsonReaderException>(() => ConfigJson.Linearize(commented));
        }

        [Test]
        public void Format_OfBrokenJson_Throws()
            => Assert.Throws<Newtonsoft.Json.JsonReaderException>(() => ConfigJson.Format("{\"a\": }"));
    }

    [TestFixture]
    [Category("PackageResources")]
    public class ResourceEditorPreferencesTests
    {
        private string _folder;
        private string _path;
        private readonly IFileSystem _fileSystem = new FileSystem();

        [SetUp]
        public void SetUp()
        {
            _folder = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "rdprefs-" + Guid.NewGuid().ToString("N"));
            _path = System.IO.Path.Combine(_folder, "sub", "resource-editor.json");
        }

        [TearDown]
        public void TearDown()
        {
            if (Directory.Exists(_folder))
            {
                Directory.Delete(_folder, true);
            }
        }

        [Test]
        public void WithoutAFile_ItIsTheDefault()
        {
            var preferences = ResourceEditorPreferences.Load(_fileSystem, _path);

            Assert.AreEqual("Consolas", preferences.FontName);
            Assert.AreEqual(10f, preferences.FontSize);
        }

        [Test]
        public void WhatWasSaved_IsLoadedAgain_AndTheFolderIsMade()
        {
            new ResourceEditorPreferences { FontName = "Cascadia Mono", FontSize = 14f }.Save(_fileSystem, _path);
            var loaded = ResourceEditorPreferences.Load(_fileSystem, _path);

            Assert.AreEqual("Cascadia Mono", loaded.FontName);
            Assert.AreEqual(14f, loaded.FontSize);
        }

        [Test]
        public void ThePreviewFont_IsSavedAndLoaded_AndAFileWithoutOneGetsTheDefault()
        {
            new ResourceEditorPreferences { PreviewFontName = "Georgia" }.Save(_fileSystem, _path);
            Assert.AreEqual("Georgia", ResourceEditorPreferences.Load(_fileSystem, _path).PreviewFontName);

            File.WriteAllText(_path, "{\"FontName\":\"Consolas\",\"FontSize\":10}");
            Assert.AreEqual("Segoe UI", ResourceEditorPreferences.Load(_fileSystem, _path).PreviewFontName);
        }

        [Test]
        public void AnUnreadableFile_FallsBackToTheDefault()
        {
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(_path));
            File.WriteAllText(_path, "not json");

            Assert.AreEqual("Consolas", ResourceEditorPreferences.Load(_fileSystem, _path).FontName);
        }

        [TestCase(1f, 6f)]
        [TestCase(500f, 48f)]
        [TestCase(12f, 12f)]
        public void TheSizeStaysInRange(float wanted, float expected)
            => Assert.AreEqual(expected, ResourceEditorPreferences.ClampSize(wanted));

        [Test]
        public void ASavedSizeOutOfRange_IsBroughtBack()
        {
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(_path));
            File.WriteAllText(_path, "{\"FontName\":\"\",\"FontSize\":900}");

            var loaded = ResourceEditorPreferences.Load(_fileSystem, _path);

            Assert.AreEqual("Consolas", loaded.FontName);
            Assert.AreEqual(48f, loaded.FontSize);
        }

        [Test]
        public void TheJsonColourDefinition_IsValid_AndXmlHasOne()
        {
            Assert.IsNotNull(Rubberduck.UI.CodeExplorer.ResourceSyntaxColors.Json);
            Assert.AreEqual("Json", Rubberduck.UI.CodeExplorer.ResourceSyntaxColors.Json.Name);
            Assert.IsNotNull(Rubberduck.UI.CodeExplorer.ResourceSyntaxColors.Xml);
        }
    }

    [TestFixture]
    [Category("PackageResources")]
    public class ResourceNamingTests
    {
        [TestCase("logo.png", "logo.png")]
        [TestCase("my logo (1).png", "my_logo__1_.png")]
        [TestCase("1st.png", "img_1st.png")]
        [TestCase("_ok.gif", "_ok.gif")]
        public void ResourceNameFor_MakesAValidRelationshipId(string file, string expected)
        {
            Assert.AreEqual(expected, ResourceEditorService.ResourceNameFor(file));
            Assert.IsTrue(PackageResourceNames.IsImageName(expected));
        }
    }
}
