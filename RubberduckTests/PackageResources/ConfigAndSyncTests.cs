using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Abstractions;
using System.IO.Compression;
using System.Linq;
using System.Text;
using NUnit.Framework;
using Rubberduck.PackageResources;
using Rubberduck.Settings;
using Rubberduck.UI.CodeExplorer.GitSync;

namespace RubberduckTests.PackageResources
{
    [TestFixture]
    [Category("PackageResources")]
    public class ConfigResourceTests
    {
        private const string Config = "{ // settings\r\n  \"sync\": { \"onViolation\": \"warn\" }\r\n}";

        private string _folder;
        private string _document;
        private PackageResourceService _service;

        [SetUp]
        public void SetUp()
        {
            _folder = Path.Combine(Path.GetTempPath(), "rd-config-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_folder);
            _document = Path.Combine(_folder, "Book.xlsm");
            PackageResourceServiceTests.WriteDocument(_document, withRibbon: true);
            _service = new PackageResourceService(new FileSystem(), Path.Combine(_folder, "journal"), Path.Combine(_folder, "backup"), () => Config);
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

        private static string Entry(string path, string name)
        {
            using (var zip = ZipFile.OpenRead(path))
            using (var reader = new StreamReader(zip.GetEntry(name).Open()))
            {
                return reader.ReadToEnd();
            }
        }

        [Test]
        public void EnsureConfig_CreatesOne_Once_AndStagesIt()
        {
            Assert.IsTrue(_service.EnsureConfig(_document));

            var config = _service.Read(_document, "rubberduck.json");
            Assert.AreEqual(Config, config.Text);
            Assert.AreEqual(PackageResourceChange.Added, config.Change);
            Assert.AreEqual(PackageResourceKind.Config, config.Kind);

            Assert.IsFalse(_service.EnsureConfig(_document), "a second call does nothing");
            _service.Discard(_document);
            Assert.IsFalse(_service.EnsureConfig(_document), "and a discarded config does not come back this session");
            Assert.IsFalse(_service.List(_document).Any(r => r.Kind == PackageResourceKind.Config));
        }

        [Test]
        public void EnsureConfig_LeavesAnExistingOneAlone_AndSkipsUnsavedDocuments()
        {
            _service.SetConfig(_document, Config, null);
            _service.TryApply(_document, out _);
            var other = new PackageResourceService(new FileSystem(), Path.Combine(_folder, "journal"), Path.Combine(_folder, "backup"), () => "{}");

            Assert.IsFalse(other.EnsureConfig(_document));
            Assert.IsFalse(other.EnsureConfig(Path.Combine(_folder, "Missing.xlsm")));
            Assert.IsFalse(other.HasPending(_document));
        }

        [Test]
        public void List_PutsTheConfigBetweenTheRibbonAndThePictures()
        {
            _service.EnsureConfig(_document);

            CollectionAssert.AreEqual(new[] { "customUI14.xml", "rubberduck.json", "logo.png" }, _service.List(_document).Select(r => r.Name).ToArray());
        }

        [Test]
        public void SetConfig_RefusesWhatIsNotAJsonObject_ButOnlyWarnsAboutUnknownSettings()
        {
            Assert.AreEqual(PackageResourceError.BadRequest, Assert.Throws<PackageResourceException>(() => _service.SetConfig(_document, "{ \"sync\": ", null)).Error);
            Assert.AreEqual(PackageResourceError.BadRequest, Assert.Throws<PackageResourceException>(() => _service.SetConfig(_document, "[1]", null)).Error);
            Assert.AreEqual(PackageResourceError.BadRequest, Assert.Throws<PackageResourceException>(() => _service.SetConfig(_document, " ", null)).Error);

            Assert.DoesNotThrow(() => _service.SetConfig(_document, "{ \"sync\": { \"nonsense\": 1 } }", null));
            CollectionAssert.IsNotEmpty(_service.CheckConfig("{ \"sync\": { \"nonsense\": 1 } }"));
            CollectionAssert.IsEmpty(_service.CheckConfig(Config));
        }

        [Test]
        public void SetConfig_NeedsTheHashOnceItExists()
        {
            _service.SetConfig(_document, Config, null);

            Assert.AreEqual(PackageResourceError.HashRequired, Assert.Throws<PackageResourceException>(() => _service.SetConfig(_document, "{}", null)).Error);
            Assert.DoesNotThrow(() => _service.SetConfig(_document, "{}", _service.Read(_document, "rubberduck.json").Hash));
        }

        [Test]
        public void RemoveImages_RefusesTheConfig()
        {
            _service.SetConfig(_document, Config, null);

            Assert.AreEqual(PackageResourceError.BadRequest, Assert.Throws<PackageResourceException>(() => _service.RemoveImages(_document, new[] { "rubberduck.json" })).Error);
        }

        [Test]
        public void TryApply_StoresTheConfigInThePackage_WithItsRelationshipAndContentType()
        {
            _service.SetConfig(_document, Config, null);

            Assert.AreEqual(PackageApplyOutcome.Applied, _service.TryApply(_document, out _));

            Assert.AreEqual(Config, Entry(_document, "customUI/rubberduck.json"));
            StringAssert.Contains("customUI/rubberduck.json", Entry(_document, "_rels/.rels"));
            StringAssert.Contains("Extension=\"json\"", Entry(_document, "[Content_Types].xml"));
            Assert.AreEqual(Config, new PackageResourceService(new FileSystem(), Path.Combine(_folder, "j2"), Path.Combine(_folder, "b2")).Read(_document, "rubberduck.json").Text);
        }

        [Test]
        public void TryApply_KeepsAnApplyingTwiceFromDuplicatingTheRelationship()
        {
            _service.SetConfig(_document, Config, null);
            _service.TryApply(_document, out _);
            _service.SetConfig(_document, "{}", _service.Read(_document, "rubberduck.json").Hash);
            _service.TryApply(_document, out _);

            var rels = Entry(_document, "_rels/.rels");
            Assert.AreEqual(1, rels.Split(new[] { "rubberduck.json" }, StringSplitOptions.None).Length - 1);
            Assert.AreEqual("{}", Entry(_document, "customUI/rubberduck.json"));
        }
    }

    [TestFixture]
    [Category("PackageResources")]
    public class ResourceSyncTests
    {
        private const string Ribbon = PackageResourceServiceTests.Ribbon;
        private static readonly byte[] Png = PackageResourceServiceTests.Png;

        private string _root;
        private string _sync;
        private string _document;
        private PackageResourceService _service;
        private GitSyncResources _sync_;
        private GitSyncState _state;
        private GitSyncScan _scan;
        private string _configPath;

        [SetUp]
        public void SetUp()
        {
            _root = Path.Combine(Path.GetTempPath(), "rd-sync-" + Guid.NewGuid().ToString("N"));
            _sync = Path.Combine(_root, "repo");
            Directory.CreateDirectory(_sync);
            _document = Path.Combine(_root, "Book.xlsm");
            PackageResourceServiceTests.WriteDocument(_document, withRibbon: true);
            _service = new PackageResourceService(new FileSystem(), Path.Combine(_root, "journal"), Path.Combine(_root, "backup"), () => "{\r\n}");
            _sync_ = new GitSyncResources(new FileSystem(), _service);
            _state = new GitSyncState();
            _configPath = Path.Combine(_sync, "rubberduck.json");
            File.WriteAllText(_configPath, "{\r\n}");
            _scan = new GitSyncScan { Folder = _sync, State = _state, Config = new ProjectConfig { Path = _configPath } };
        }

        [TearDown]
        public void TearDown()
        {
            try
            {
                Directory.Delete(_root, true);
            }
            catch (IOException)
            {
            }
        }

        private SyncPlan Plan()
        {
            var plan = new SyncPlan();
            _sync_.Plan(_document, _sync, _configPath, _state, new GitIgnoreSet(new KeyValuePair<string, GitIgnore>[0]), plan);
            return plan;
        }

        private static SyncItem Item(SyncPlan plan, string name) => plan.Items.Single(i => i.Name == name);

        private GitSyncResult Apply(SyncPlan plan, params (string Name, SyncAction Action)[] choices)
        {
            var result = new GitSyncResult();
            foreach (var (name, action) in choices)
            {
                Item(plan, name).Chosen = action;
            }

            _sync_.Apply(_document, _scan, plan.Items, path => { }, result);
            return result;
        }

        [Test]
        public void FirstScan_ListsWhatOnlyTheDocumentHas_AndSuggestsExporting()
        {
            var plan = Plan();

            CollectionAssert.AreEquivalent(new[] { "Resources/customUI14.xml", "Resources/images/logo.png", "rubberduck.json" }, plan.Items.Select(i => i.Name).ToArray());
            Assert.IsTrue(plan.Items.All(i => i.IsResource));
            Assert.IsTrue(plan.Items.Where(i => i.Name != "rubberduck.json").All(i => i.Status == SyncStatus.ModuleOnly && i.Suggested == SyncAction.Export));
            Assert.AreEqual(SyncStatus.FileOnly, Item(plan, "rubberduck.json").Status, "the folder has a config and the document none");
        }

        [Test]
        public void Export_WritesFilesIntoTheResourcesFolder_AndThenEverythingIsInSync()
        {
            var plan = Plan();
            var result = Apply(plan, ("Resources/customUI14.xml", SyncAction.Export), ("Resources/images/logo.png", SyncAction.Export));

            Assert.AreEqual(2, result.Exported);
            CollectionAssert.IsEmpty(result.Errors);
            StringAssert.Contains("<ribbon", File.ReadAllText(Path.Combine(_sync, "Resources", "customUI14.xml")));
            CollectionAssert.AreEqual(Png, File.ReadAllBytes(Path.Combine(_sync, "Resources", "images", "logo.png")));

            var again = Plan();
            Assert.IsFalse(again.Items.Any(i => i.Name.StartsWith("Resources/")));
            Assert.AreEqual(2, again.InSyncBaselines.Count(b => b.Key.StartsWith("Resources/")));
        }

        [Test]
        public void ChangeInTheFolder_IsSuggestedAsAnImport_AndStagedInTheDocument()
        {
            Apply(Plan(), ("Resources/customUI14.xml", SyncAction.Export), ("Resources/images/logo.png", SyncAction.Export));
            var path = Path.Combine(_sync, "Resources", "customUI14.xml");
            var edited = File.ReadAllText(path).Replace("<ribbon/>", "<ribbon><tabs/></ribbon>");
            File.WriteAllText(path, edited);
            File.WriteAllBytes(Path.Combine(_sync, "Resources", "images", "other.png"), Png.Concat(new byte[] { 7 }).ToArray());

            var plan = Plan();
            Assert.AreEqual(SyncStatus.FileChanged, Item(plan, "Resources/customUI14.xml").Status);
            Assert.AreEqual(SyncAction.Import, Item(plan, "Resources/customUI14.xml").Suggested);
            Assert.AreEqual(SyncStatus.FileOnly, Item(plan, "Resources/images/other.png").Status);
            var result = Apply(plan, ("Resources/customUI14.xml", SyncAction.Import), ("Resources/images/other.png", SyncAction.Import));

            Assert.AreEqual(2, result.Imported, string.Join("; ", result.Errors));
            Assert.AreEqual(edited, _service.Read(_document, "customUI14.xml").Text);
            Assert.AreEqual(PackageResourceChange.Added, _service.List(_document).Single(r => r.Name == "other.png").Change);
            Assert.IsTrue(_service.HasPending(_document), "the document is open, so the import is staged");
            Assert.IsFalse(Plan().Items.Any(i => i.Name.StartsWith("Resources/")), "and the next scan already sees both sides alike");
        }

        [Test]
        public void ChangeInTheDocument_IsSuggestedAsAnExport()
        {
            Apply(Plan(), ("Resources/customUI14.xml", SyncAction.Export), ("Resources/images/logo.png", SyncAction.Export));
            _service.SetXml(_document, "customUI14.xml", Ribbon.Replace("<ribbon/>", "<ribbon><tabs/></ribbon>"), _service.Read(_document, "customUI14.xml").Hash);

            var item = Item(Plan(), "Resources/customUI14.xml");

            Assert.AreEqual(SyncStatus.ModuleChanged, item.Status);
            Assert.AreEqual(SyncAction.Export, item.Suggested);
            Assert.IsTrue(item.CanCompare, "text can be compared and merged line by line");
        }

        [Test]
        public void ChangeOnBothSides_CanBeMerged_AndBothEndUpWithTheMergedText()
        {
            Apply(Plan(), ("Resources/customUI14.xml", SyncAction.Export), ("Resources/images/logo.png", SyncAction.Export));
            _service.SetXml(_document, "customUI14.xml", Ribbon.Replace("<ribbon/>", "<ribbon><tabs/></ribbon>"), _service.Read(_document, "customUI14.xml").Hash);
            File.WriteAllText(Path.Combine(_sync, "Resources", "customUI14.xml"), Ribbon.Replace("<ribbon/>", "<ribbon><qat/></ribbon>"));

            var plan = Plan();
            var item = Item(plan, "Resources/customUI14.xml");
            Assert.AreEqual(SyncStatus.BothChanged, item.Status);
            item.MergedText = Ribbon.Replace("<ribbon/>", "<ribbon><tabs/><qat/></ribbon>");
            var result = Apply(plan, ("Resources/customUI14.xml", SyncAction.Merge));

            Assert.AreEqual(1, result.Merged, string.Join("; ", result.Errors));
            StringAssert.Contains("<tabs/><qat/>", _service.Read(_document, "customUI14.xml").Text);
            StringAssert.Contains("<tabs/><qat/>", File.ReadAllText(Path.Combine(_sync, "Resources", "customUI14.xml")));
        }

        [Test]
        public void Config_IsSyncedWithTheFoldersRubberduckJson()
        {
            var plan = Plan();
            Assert.AreEqual(SyncAction.Import, Item(plan, "rubberduck.json").Suggested);

            File.WriteAllText(_configPath, "{ \"sync\": { \"onViolation\": \"block\" } }");
            plan = Plan();
            var result = Apply(plan, ("rubberduck.json", SyncAction.Import));

            Assert.AreEqual(1, result.Imported, string.Join("; ", result.Errors));
            StringAssert.Contains("block", _service.Read(_document, "rubberduck.json").Text);

            _service.SetConfig(_document, "{ \"sync\": { \"onViolation\": \"warn\" } }", _service.Read(_document, "rubberduck.json").Hash);
            Assert.AreEqual(SyncAction.Export, Item(Plan(), "rubberduck.json").Suggested);
        }

        [Test]
        public void ADifferenceInLineEndingsOrTrailingBlankLines_IsNotAChange()
        {
            Apply(Plan(), ("Resources/customUI14.xml", SyncAction.Export), ("Resources/images/logo.png", SyncAction.Export));
            var path = Path.Combine(_sync, "Resources", "customUI14.xml");
            File.WriteAllText(path, File.ReadAllText(path).Replace("\r\n", "\n") + "\n\n");

            Assert.IsFalse(Plan().Items.Any(i => i.Name == "Resources/customUI14.xml"));
        }

        [Test]
        public void IgnoredFiles_AreLeftOut()
        {
            var plan = new SyncPlan();
            var ignore = new GitIgnoreSet(new[] { new KeyValuePair<string, GitIgnore>(string.Empty, new GitIgnore(new[] { "/Resources/images/" })) });

            _sync_.Plan(_document, _sync, _configPath, _state, ignore, plan);

            Assert.IsFalse(plan.Items.Any(i => i.Name.Contains("images")));
            Assert.Contains("Resources/images/logo.png", plan.Ignored);
        }

        [Test]
        public void AnImportThatTheDocumentRefuses_IsReportedAndTheRestGoesOn()
        {
            File.WriteAllBytes(Path.Combine(_sync, "Resources", "images", "fake.png").EnsureFolder(), new byte[] { 1, 2, 3 });
            var plan = Plan();

            var result = Apply(plan, ("Resources/images/fake.png", SyncAction.Import), ("Resources/customUI14.xml", SyncAction.Export));

            Assert.AreEqual(1, result.Errors.Count);
            StringAssert.Contains("fake.png", result.Errors[0]);
            Assert.AreEqual(1, result.Exported);
        }

        [Test]
        public void ADocumentWithoutRibbonFolder_IsNotSynced()
        {
            Assert.IsFalse(_sync_.AppliesTo(Path.Combine(_root, "Old.xls")));
            Assert.IsTrue(_sync_.AppliesTo(_document));
        }
    }

    internal static class PathExtensions
    {
        public static string EnsureFolder(this string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            return path;
        }
    }
}
