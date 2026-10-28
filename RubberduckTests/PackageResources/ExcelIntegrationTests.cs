using System;
using System.Diagnostics;
using System.IO;
using System.IO.Abstractions;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Moq;
using NUnit.Framework;
using Rubberduck.PackageResources;
using Rubberduck.VBEditor.Events;
using Rubberduck.Settings;
using Rubberduck.UI.CodeExplorer.GitSync;

namespace RubberduckTests.PackageResources
{
    /// <summary>
    /// The resources against a real Excel: does it open what Rubberduck writes, keep it when it saves, and lock the file the way the staging assumes?
    /// Slow, and it starts and stops Excel (visibly), so it only runs when asked for:
    /// <c>set RUBBERDUCK_EXCEL_TESTS=1</c> and select the category <c>Excel</c>. Needs Excel with "Trust access to the VBA project object model" switched on.
    /// </summary>
    [TestFixture]
    [Category("Excel")]
    [NonParallelizable]
    public class ExcelIntegrationTests
    {
        private const string OneByOnePng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

        private const string RibbonXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n" +
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"OnRibbonLoad\">\r\n" +
            "  <ribbon><tabs><tab id=\"rdTab\" label=\"RD Test\"><group id=\"rdGroup\" label=\"Group\">\r\n" +
            "    <button id=\"rdButton\" label=\"Run\" size=\"large\" image=\"logo\" onAction=\"OnButton\" />\r\n" +
            "  </group></tab></tabs></ribbon>\r\n" +
            "</customUI>";

        private const string ConfigJsonText = "{\r\n  // from the Excel test\r\n  \"sync\": { \"onViolation\": \"warn\", \"ignoreCaseOutsideStrings\": true }\r\n}";

        private const string Code =
            "Public Sub OnRibbonLoad(ByVal ribbon As IRibbonUI)\r\n" +
            "    ThisWorkbook.Worksheets(1).Range(\"B1\").Value = \"ribbon loaded\"\r\n" +
            "End Sub\r\n" +
            "Public Sub OnButton(ByVal control As IRibbonControl)\r\n" +
            "End Sub\r\n" +
            "Public Function Answer() As Long\r\n" +
            "    Answer = 42\r\n" +
            "End Function\r\n";

        private static readonly byte[] Png = Convert.FromBase64String(OneByOnePng);

        private string _folder;
        private dynamic _excel;
        private int _excelProcess;
        private PackageResourceService _service;

        [SetUp]
        public void SetUp()
        {
            Assume.That(Environment.GetEnvironmentVariable("RUBBERDUCK_EXCEL_TESTS"), Is.EqualTo("1"), "Set RUBBERDUCK_EXCEL_TESTS=1 to run the tests that start Excel.");
            Assume.That(Type.GetTypeFromProgID("Excel.Application"), Is.Not.Null, "Excel is not installed.");

            _folder = Path.Combine(Path.GetTempPath(), "rd-excel-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_folder);
            _service = new PackageResourceService(new FileSystem(), Path.Combine(_folder, "journal"), Path.Combine(_folder, "backup"), () => ConfigJsonText);

            _excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            _excel.DisplayAlerts = false;
            _excel.Visible = true;
            _excel.AutomationSecurity = 1; // macros enabled: the ribbon callbacks are VBA
            GetWindowThreadProcessId(new IntPtr((int)_excel.Hwnd), out var process);
            _excelProcess = (int)process;
        }

        [TearDown]
        public void TearDown()
        {
            if (_excel != null)
            {
                try
                {
                    foreach (var workbook in ((System.Collections.IEnumerable)_excel.Workbooks).Cast<dynamic>().ToList())
                    {
                        workbook.Close(false);
                    }

                    _excel.Quit();
                }
                catch (COMException)
                {
                }

                Marshal.FinalReleaseComObject(_excel);
                _excel = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //Excel can linger for a moment; a test must not leave one behind.
                try
                {
                    var process = Process.GetProcessById(_excelProcess);
                    if (!process.WaitForExit(8000))
                    {
                        process.Kill();
                    }
                }
                catch (ArgumentException)
                {
                }
            }

            try
            {
                if (_folder != null)
                {
                    Directory.Delete(_folder, true);
                }
            }
            catch (IOException)
            {
            }
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

        #region Helpers

        //A real macro-enabled workbook with some VBA in it, made by Excel itself, left closed.
        private string NewWorkbook(string name = "Book.xlsm")
        {
            var path = Path.Combine(_folder, name);
            dynamic workbook = _excel.Workbooks.Add();
            dynamic module = workbook.VBProject.VBComponents.Add(1);
            module.Name = "modRibbon";
            module.CodeModule.AddFromString(Code);
            workbook.SaveAs(path, 52);
            workbook.Close(false);
            return path;
        }

        //Excel holds its workbook open; reading it is allowed as long as the reader does not insist on exclusivity.
        private static byte[] ReadShared(string path)
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var copy = new MemoryStream())
            {
                stream.CopyTo(copy);
                return copy.ToArray();
            }
        }

        private dynamic Open(string path) => _excel.Workbooks.Open(path);

        private static void CloseWorkbook(dynamic workbook, bool save = false) => workbook.Close(save);

        private static object Cell(dynamic workbook, string address) => workbook.Worksheets[1].Range(address).Value2;

        private static bool WaitFor(Func<bool> condition, int seconds = 20)
        {
            var until = DateTime.UtcNow.AddSeconds(seconds);
            while (DateTime.UtcNow < until)
            {
                try
                {
                    if (condition())
                    {
                        return true;
                    }
                }
                catch (COMException)
                {
                    //Excel is busy loading; ask again.
                }

                Thread.Sleep(250);
            }

            return false;
        }

        private static string Entry(string path, string name)
        {
            using (var zip = ZipFile.OpenRead(path))
            {
                var entry = zip.GetEntry(name);
                if (entry == null)
                {
                    return null;
                }

                using (var reader = new StreamReader(entry.Open()))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        private void StageEverything(string path)
        {
            _service.SetXml(path, "customUI14.xml", RibbonXml, null);
            _service.SetImage(path, "logo.png", Png, null);
            _service.SetConfig(path, ConfigJsonText, null);
        }

        private void AssertEverythingIsThere(string path)
        {
            var names = _service.List(path).Select(r => r.Name).ToList();
            CollectionAssert.IsSubsetOf(new[] { "customUI14.xml", "rubberduck.json", "logo.png" }, names);
            Assert.AreEqual(ConfigJsonText, _service.Read(path, "rubberduck.json").Text);
            CollectionAssert.AreEqual(Png, _service.Read(path, "logo.png").Bytes);
            StringAssert.Contains("onLoad=\"OnRibbonLoad\"", _service.Read(path, "customUI14.xml").Text);
        }

        #endregion

        [Test]
        public void ExcelOwnWorkbook_IsRecognisedAsHavingNoResourcesYet()
        {
            var path = NewWorkbook();

            Assert.IsTrue(_service.IsSupported(path));
            CollectionAssert.IsEmpty(_service.List(path));
        }

        [Test]
        public void WhatIsAppliedToAClosedWorkbook_IsLoadedByExcel_RibbonIncluded_AndVbaIsUntouched()
        {
            var path = NewWorkbook();
            StageEverything(path);

            Assert.AreEqual(PackageApplyOutcome.Applied, _service.TryApply(path, out var message), message);

            var workbook = Open(path);
            try
            {
                Assert.AreEqual("modRibbon", (string)workbook.VBProject.VBComponents.Item("modRibbon").Name);
                Assert.AreEqual(42, (int)_excel.Run("'" + workbook.Name + "'!Answer"), "the VBA survived the patching");

                //The ribbon's onLoad is a VBA callback: it only runs if Excel accepted customUI14.xml, its namespace and the parts around it.
                Assert.IsTrue(WaitFor(() => (string)Cell(workbook, "B1") == "ribbon loaded"), "Excel did not load the ribbon");
            }
            finally
            {
                CloseWorkbook(workbook);
            }
        }

        [Test]
        public void ExcelKeepsTheRibbonPicturesAndConfigWhenItSavesTheWorkbook()
        {
            var path = NewWorkbook();
            StageEverything(path);
            _service.TryApply(path, out _);

            var workbook = Open(path);
            workbook.Worksheets[1].Range("A1").Value2 = "saved by excel";
            workbook.Save();
            CloseWorkbook(workbook);

            AssertEverythingIsThere(path);
            Assert.IsNotNull(Entry(path, "customUI/rubberduck.json"), "Excel dropped the config part");
            StringAssert.Contains("rubberduck.json", Entry(path, "_rels/.rels"));
        }

        [Test]
        public void WhileExcelHasTheWorkbookOpen_TheFileIsLocked_ChangesAreStaged_AndNothingIsLost()
        {
            var path = NewWorkbook();
            var workbook = Open(path);
            var before = ReadShared(path);

            StageEverything(path);

            //Reading the saved file works while Excel holds it; writing does not.
            Assert.AreEqual(PackageApplyOutcome.Busy, _service.TryApply(path, out _));
            CollectionAssert.AreEqual(before, ReadShared(path));
            AssertEverythingIsThere(path);
            Assert.IsTrue(_service.HasPending(path));
            Assert.IsFalse(File.Exists(path + ".rd-resources.tmp"));

            //Excel is still happy with its workbook.
            workbook.Worksheets[1].Range("A1").Value2 = "typed meanwhile";
            workbook.Save();
            CloseWorkbook(workbook);

            Assert.AreEqual(PackageApplyOutcome.Applied, _service.TryApply(path, out var message), message);
            AssertEverythingIsThere(path);

            var reopened = Open(path);
            try
            {
                Assert.AreEqual("typed meanwhile", (string)Cell(reopened, "A1"), "what Excel saved in the meantime is still there");
                Assert.AreEqual(42, (int)_excel.Run("'" + reopened.Name + "'!Answer"));
            }
            finally
            {
                CloseWorkbook(reopened);
            }
        }

        [Test]
        public void TheApplier_WaitsForExcelToClose_AndThenWritesTheStagedChanges()
        {
            var path = NewWorkbook();
            var workbook = Open(path);
            StageEverything(path);

            var applier = new PackageResourceApplier(_service, new Mock<IVbeEvents>().Object);
            var pass = Task.Run(() => applier.ApplyPending(TimeSpan.FromSeconds(60), CancellationToken.None));

            Thread.Sleep(3000);
            Assert.IsTrue(_service.HasPending(path), "still staged while the workbook is open");

            CloseWorkbook(workbook);
            Assert.IsTrue(pass.Wait(TimeSpan.FromSeconds(30)), "the applier did not finish after the workbook was closed");

            Assert.IsFalse(_service.HasPending(path));
            AssertEverythingIsThere(path);
            Assert.IsNotNull(Entry(path, "customUI/images/logo.png"));
            Assert.AreEqual(1, Directory.GetFiles(Path.Combine(_folder, "backup"), "Book.xlsm", SearchOption.AllDirectories).Length, "the original is kept");
        }

        [Test]
        public void ReplacingAndRemovingPictures_SurvivesAnExcelRoundTrip()
        {
            var path = NewWorkbook();
            StageEverything(path);
            _service.SetImage(path, "extra.png", Png, null);
            _service.TryApply(path, out _);

            _service.RemoveImages(path, new[] { "extra.png" });
            _service.SetXml(path, "customUI14.xml", RibbonXml.Replace("label=\"Run\"", "label=\"Go\""), _service.Read(path, "customUI14.xml").Hash);
            _service.TryApply(path, out _);

            var workbook = Open(path);
            workbook.Save();
            CloseWorkbook(workbook);

            var names = _service.List(path).Select(r => r.Name).ToList();
            CollectionAssert.DoesNotContain(names, "extra.png");
            CollectionAssert.Contains(names, "logo.png");
            StringAssert.Contains("label=\"Go\"", _service.Read(path, "customUI14.xml").Text);
            StringAssert.DoesNotContain("extra.png", Entry(path, "customUI/_rels/customUI14.xml.rels"));
        }

        [Test]
        public void AWorkbookWithoutResources_GetsAConfig_ExcelKeeps_AndSyncSeesIt()
        {
            var path = NewWorkbook();

            Assert.IsTrue(_service.EnsureConfig(path));
            _service.TryApply(path, out _);

            var workbook = Open(path);
            workbook.Worksheets[1].Range("A1").Value2 = 1;
            workbook.Save();
            CloseWorkbook(workbook);

            Assert.AreEqual(ConfigJsonText, _service.Read(path, "rubberduck.json").Text);

            //Sync against an empty repository folder: the config is the one thing to export.
            var repo = Path.Combine(_folder, "repo");
            Directory.CreateDirectory(repo);
            var configPath = Path.Combine(repo, "rubberduck.json");
            File.WriteAllText(configPath, ConfigJsonText);
            var sync = new GitSyncResources(new FileSystem(), _service);
            var plan = new SyncPlan();
            sync.Plan(path, repo, configPath, new GitSyncState(), new GitIgnoreSet(new System.Collections.Generic.KeyValuePair<string, GitIgnore>[0]), plan);

            Assert.AreEqual(1, plan.InSyncCount, "the document's config and the repository's are the same text");
            CollectionAssert.IsEmpty(plan.Items);
        }

        [Test]
        public void SyncImportsARibbonAndPicturesFromAFolder_IntoAnExcelWorkbook_AndExcelLoadsThem()
        {
            var path = NewWorkbook();
            var repo = Path.Combine(_folder, "repo");
            Directory.CreateDirectory(Path.Combine(repo, "Resources", "images"));
            File.WriteAllText(Path.Combine(repo, "Resources", "customUI14.xml"), RibbonXml);
            File.WriteAllBytes(Path.Combine(repo, "Resources", "images", "logo.png"), Png);
            var configPath = Path.Combine(repo, "rubberduck.json");
            File.WriteAllText(configPath, ConfigJsonText);

            var sync = new GitSyncResources(new FileSystem(), _service);
            var state = new GitSyncState();
            var plan = new SyncPlan();
            sync.Plan(path, repo, configPath, state, new GitIgnoreSet(new System.Collections.Generic.KeyValuePair<string, GitIgnore>[0]), plan);
            Assert.AreEqual(3, plan.Items.Count);
            Assert.IsTrue(plan.Items.All(i => i.Status == SyncStatus.FileOnly && i.Suggested == SyncAction.Import));
            foreach (var item in plan.Items)
            {
                item.Chosen = SyncAction.Import;
            }

            var result = new GitSyncResult();
            sync.Apply(path, new GitSyncScan { Folder = repo, State = state, Config = new ProjectConfig { Path = configPath } }, plan.Items, p => { }, result);
            CollectionAssert.IsEmpty(result.Errors);
            Assert.AreEqual(3, result.Imported);

            Assert.AreEqual(PackageApplyOutcome.Applied, _service.TryApply(path, out var message), message);
            var workbook = Open(path);
            try
            {
                Assert.IsTrue(WaitFor(() => (string)Cell(workbook, "B1") == "ribbon loaded"), "Excel did not load the imported ribbon");
            }
            finally
            {
                CloseWorkbook(workbook);
            }

            //Nothing left to sync.
            var again = new SyncPlan();
            sync.Plan(path, repo, configPath, state, new GitIgnoreSet(new System.Collections.Generic.KeyValuePair<string, GitIgnore>[0]), again);
            CollectionAssert.IsEmpty(again.Items);
        }
    }
}
