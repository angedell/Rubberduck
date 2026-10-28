using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Abstractions;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Moq;
using NUnit.Framework;
using Rubberduck.PackageResources;
using Rubberduck.Ship;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Path = System.IO.Path;

namespace RubberduckTests.Ship
{
    /// <summary>
    /// Ship against a real Excel: a real workbook is built, shipped to .xlsm and .xlam by the real <see cref="ShipArtifactWriter"/>, and the artifact is reopened.
    /// Slow, and it starts Excel (hidden), so it only runs when asked for: <c>set RUBBERDUCK_EXCEL_TESTS=1</c> and select the category <c>Excel</c>.
    /// Needs Excel with "Trust access to the VBA project object model". Every Excel it starts is a new process; it never uses or closes one that was already running.
    /// </summary>
    [TestFixture]
    [Category("Excel")]
    [NonParallelizable]
    public class ShipExcelIntegrationTests
    {
        private const string ConfigJsonText = "{ \"sync\": { \"onViolation\": \"warn\" } }";

        private const string MathModule =
            "Attribute VB_Name = \"modMath\"\r\n" +
            "Option Explicit\r\n" +
            "' The answer to everything\r\n" +
            "Public Function Answer() As Long\r\n" +
            "Attribute Answer.VB_Description = \"The answer\"\r\n" +
            "    Answer = 42 ' really\r\n" +
            "End Function\r\n" +
            "Private Function Secret() As Long\r\n" +
            "Attribute Secret.VB_Description = \"Not for Excel\"\r\n" +
            "    Secret = 1\r\n" +
            "End Function\r\n" +
            "Public Function Quote() As String\r\n" +
            "    Quote = \"it's ' not a comment\" ' but this is\r\n" +
            "End Function\r\n" +
            "#If Win64 Then\r\n" +
            "    ' sixty-four bit comment\r\n" +
            "#Else\r\n" +
            "    ' thirty-two bit comment\r\n" +
            "#End If\r\n" +
            "Public Function ThingHello() As String\r\n" +
            "    Dim t As New Thing\r\n" +
            "    ThingHello = t.Hello\r\n" +
            "End Function\r\n";

        private const string ThingClass =
            "VERSION 1.0 CLASS\r\n" +
            "BEGIN\r\n" +
            "  MultiUse = -1  'True\r\n" +
            "END\r\n" +
            "Attribute VB_Name = \"Thing\"\r\n" +
            "Attribute VB_GlobalNameSpace = False\r\n" +
            "Attribute VB_Creatable = False\r\n" +
            "Attribute VB_PredeclaredId = False\r\n" +
            "Attribute VB_Exposed = False\r\n" +
            "Attribute VB_Ext_KEY = \"Rubberduck\" ,\"Predeclared\"\r\n" +
            "Option Explicit\r\n" +
            "' about Thing\r\n" +
            "Public Function Hello() As String\r\n" +
            "    Hello = \"hello\" ' hi\r\n" +
            "End Function\r\n";

        private const string TestModuleCode =
            "Attribute VB_Name = \"TestMath\"\r\n'@TestModule\r\nOption Explicit\r\n'@TestMethod\r\nPublic Sub AnswerIs42()\r\nEnd Sub\r\n";

        private const string NoteCode =
            "Attribute VB_Name = \"Note_1\"\r\n'@Notes(\"README.md\")\r\n' # Notes\r\n' words\r\n";

        private string _folder;
        private readonly List<ExcelSession> _sessions = new List<ExcelSession>();

        [SetUp]
        public void SetUp()
        {
            Assume.That(Environment.GetEnvironmentVariable("RUBBERDUCK_EXCEL_TESTS"), Is.EqualTo("1"), "Set RUBBERDUCK_EXCEL_TESTS=1 to run the tests that start Excel.");
            Assume.That(Type.GetTypeFromProgID("Excel.Application"), Is.Not.Null, "Excel is not installed.");
            _folder = Path.Combine(Path.GetTempPath(), "rd-ship-excel-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_folder);
        }

        [TearDown]
        public void TearDown()
        {
            foreach (var session in _sessions)
            {
                session.Dispose();
            }

            _sessions.Clear();

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

        /// <summary>A new Excel process of the test's own. Never the one that was running before.</summary>
        private class ExcelSession : IDisposable
        {
            private readonly int _process;
            public dynamic Excel { get; }

            public ExcelSession(int securityLevel)
            {
                var before = new HashSet<int>(Process.GetProcessesByName("EXCEL").Select(p => p.Id));
                Excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                GetWindowThreadProcessId(new IntPtr((int)Excel.Hwnd), out var process);
                _process = (int)process;
                if (before.Contains(_process))
                {
                    Marshal.FinalReleaseComObject((object)Excel);
                    Assert.Inconclusive("COM handed out an Excel that was already running; the test does not touch it.");
                }

                Excel.DisplayAlerts = false;
                Excel.Visible = false;
                Excel.AutomationSecurity = securityLevel;
            }

            public void Dispose()
            {
                try
                {
                    foreach (var workbook in ((System.Collections.IEnumerable)Excel.Workbooks).Cast<dynamic>().ToList())
                    {
                        workbook.Close(false);
                    }

                    Excel.Quit();
                }
                catch (COMException)
                {
                }

                Marshal.FinalReleaseComObject((object)Excel);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                try
                {
                    var process = Process.GetProcessById(_process);
                    if (!process.WaitForExit(8000))
                    {
                        process.Kill();
                    }
                }
                catch (ArgumentException)
                {
                }
            }

            [DllImport("user32.dll")]
            private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
        }

        private ExcelSession NewExcel(int securityLevel = 1)
        {
            var session = new ExcelSession(securityLevel);
            _sessions.Add(session);
            return session;
        }

        private class Built
        {
            public string Workbook { get; set; }
            public List<(string Name, ComponentType Type, string Text, string File)> Components { get; } = new List<(string, ComponentType, string, string)>();
        }

        //A real macro-enabled workbook with a standard module, a class, a test module, a note, a sheet and the document module, and a rubberduck.json in its package.
        private Built BuildWorkbook()
        {
            var built = new Built { Workbook = Path.Combine(_folder, "Book.xlsm") };
            var modules = Path.Combine(_folder, "source");
            Directory.CreateDirectory(modules);
            File.WriteAllText(Path.Combine(modules, "modMath.bas"), MathModule, Encoding.Default);
            File.WriteAllText(Path.Combine(modules, "Thing.cls"), ThingClass, Encoding.Default);
            File.WriteAllText(Path.Combine(modules, "TestMath.bas"), TestModuleCode, Encoding.Default);
            File.WriteAllText(Path.Combine(modules, "Note_1.bas"), NoteCode, Encoding.Default);

            var session = NewExcel();
            dynamic workbook = session.Excel.Workbooks.Add();
            dynamic project = workbook.VBProject;
            foreach (var file in new[] { "modMath.bas", "Thing.cls", "TestMath.bas", "Note_1.bas" })
            {
                project.VBComponents.Import(Path.Combine(modules, file));
            }

            dynamic thisWorkbook = project.VBComponents.Item("ThisWorkbook").CodeModule;
            //The module may already say Option Explicit (Require Variable Declaration): replace, never append.
            int existingLines = thisWorkbook.CountOfLines;
            if (existingLines > 0)
            {
                thisWorkbook.DeleteLines(1, existingLines);
            }

            thisWorkbook.AddFromString("' about the workbook\r\nOption Explicit\r\n");
            workbook.SaveAs(built.Workbook, 52);

            var exports = Path.Combine(_folder, "exports");
            Directory.CreateDirectory(exports);
            foreach (dynamic component in ((System.Collections.IEnumerable)project.VBComponents).Cast<object>().ToList())
            {
                string name = component.Name;
                var type = (ComponentType)(int)component.Type;
                var extension = ShipTestExtension(type);
                var file = Path.Combine(exports, name + extension);
                component.Export(file);
                built.Components.Add((name, type, File.ReadAllText(file, Encoding.Default), file));
            }

            workbook.Close(false);

            //The document gets the config part the way Rubberduck stages it.
            var resources = new PackageResourceService(new FileSystem(), Path.Combine(_folder, "journal"), Path.Combine(_folder, "backup"), () => ConfigJsonText);
            resources.SetConfig(built.Workbook, ConfigJsonText, null);
            Assert.AreEqual(PackageApplyOutcome.Applied, resources.TryApply(built.Workbook, out var message), message);
            return built;
        }

        private static string ShipTestExtension(ComponentType type)
            => type == ComponentType.StandardModule ? ".bas" : type == ComponentType.UserForm ? ".frm" : ".cls";

        //The working project as Rubberduck sees it, read from the workbook that Excel built.
        private Mock<IVBProject> ProjectOf(Built built)
        {
            var project = new MockVbeBuilder().ProjectBuilder("Book", built.Workbook, ProjectProtection.Unprotected);
            foreach (var component in built.Components)
            {
                project.AddComponent(component.Name, component.Type, component.Text);
            }

            var mock = project.Build();
            mock.SetupGet(p => p.IsSaved).Returns(true);
            mock.SetupGet(p => p.ProjectId).Returns("project-1");
            var host = new Mock<IHostApplication>();
            host.SetupGet(h => h.ApplicationName).Returns("Excel");
            Mock.Get(mock.Object.VBE).Setup(v => v.HostApplication()).Returns(host.Object);
            foreach (var component in project.MockComponents)
            {
                var source = built.Components.Single(c => c.Name == component.Object.Name).File;
                component.Setup(c => c.Export(It.IsAny<string>())).Callback<string>(path => File.Copy(source, path, true));
            }

            return mock;
        }

        private static string HashOf(string path)
        {
            using (var sha = SHA256.Create())
            {
                return Convert.ToBase64String(sha.ComputeHash(File.ReadAllBytes(path)));
            }
        }

        private ShipService Service()
            => new ShipService(new FileSystem(), new PackageResourceService(new FileSystem(), Path.Combine(_folder, "journal2"), Path.Combine(_folder, "backup2")),
                ShipTestParser.Create(), new ShipArtifactWriter(new FileSystem(), new ShipModuleText(ShipTestParser.Create())), new ShipEnvironment(),
                Path.Combine(_folder, "unattended.json"), Path.Combine(_folder, "temp"));

        private static string CodeOf(dynamic component)
        {
            dynamic code = component.CodeModule;
            int count = code.CountOfLines;
            return count == 0 ? string.Empty : (string)code.Lines(1, count);
        }

        [TestCase(ShipFormat.Xlsm, ".xlsm")]
        [TestCase(ShipFormat.Xlam, ".xlam")]
        public void AShippedWorkbookIsCleanAndStillRuns(ShipFormat format, string extension)
        {
            var built = BuildWorkbook();
            var before = HashOf(built.Workbook);
            var output = Path.Combine(_folder, "out");
            Directory.CreateDirectory(output);
            var settings = new ShipSettings { Format = format, OutputFolder = output, FileName = "Tool" };

            var report = Service().Ship(ProjectOf(built).Object, settings, ShipOnExisting.Ask, null);

            var artifact = Path.Combine(output, "Tool" + extension);
            Assert.AreEqual(artifact, report.OutputPath);
            Assert.IsTrue(File.Exists(artifact));
            Assert.AreEqual(before, HashOf(built.Workbook), "The workbook that was shipped is the same file as before.");
            CollectionAssert.IsEmpty(Directory.GetFiles(output).Where(f => f != artifact), "Nothing else is left in the output folder.");
            Assert.IsEmpty(report.Unparsed, string.Join(" | ", report.Unparsed));

            //The package: no config part and no relationship to it, the file is a normal zip.
            using (var zip = ZipFile.OpenRead(artifact))
            {
                CollectionAssert.DoesNotContain(zip.Entries.Select(e => e.FullName), "customUI/rubberduck.json");
                using (var reader = new StreamReader(zip.GetEntry("_rels/.rels").Open()))
                {
                    StringAssert.DoesNotContain("rubberduck", reader.ReadToEnd());
                }
            }

            //The artifact reopened, with macros allowed, in an Excel of its own.
            var session = NewExcel(securityLevel: 1);
            dynamic opened = session.Excel.Workbooks.Open(artifact);
            dynamic project = opened.VBProject;

            var names = ((System.Collections.IEnumerable)project.VBComponents).Cast<dynamic>().Select(c => (string)c.Name).ToList();
            CollectionAssert.DoesNotContain(names, "Note_1", "The note is gone.");
            CollectionAssert.DoesNotContain(names, "TestMath", "The test module is gone.");
            CollectionAssert.IsSupersetOf(names, new[] { "modMath", "Thing", "ThisWorkbook" });

            foreach (dynamic component in ((System.Collections.IEnumerable)project.VBComponents).Cast<object>().ToList())
            {
                var code = CodeOf(component);
                var withoutLiterals = code.Replace("\"it's ' not a comment\"", string.Empty);
                StringAssert.DoesNotContain("'", withoutLiterals, "No comment is left in " + (string)component.Name);
                StringAssert.DoesNotContain("Rem ", withoutLiterals);
            }

            StringAssert.Contains("\"it's ' not a comment\"", CodeOf(project.VBComponents.Item("modMath")), "A string with a comment mark in it is not touched.");

            //Attributes: the description of a public procedure stays, the Rubberduck key does not.
            var exports = Path.Combine(_folder, "artifact-exports");
            Directory.CreateDirectory(exports);
            project.VBComponents.Item("modMath").Export(Path.Combine(exports, "modMath.bas"));
            project.VBComponents.Item("Thing").Export(Path.Combine(exports, "Thing.cls"));
            var mathText = File.ReadAllText(Path.Combine(exports, "modMath.bas"), Encoding.Default);
            var thingText = File.ReadAllText(Path.Combine(exports, "Thing.cls"), Encoding.Default);
            StringAssert.Contains("Answer.VB_Description", mathText);
            StringAssert.DoesNotContain("Secret.VB_Description", mathText);
            StringAssert.DoesNotContain("VB_Ext_KEY", thingText);
            StringAssert.Contains("VB_Name = \"Thing\"", thingText);

            //And it runs.
            var prefix = "'" + Path.GetFileName(artifact) + "'!";
            Assert.AreEqual(42, (int)session.Excel.Run(prefix + "Answer"));
            Assert.AreEqual("it's ' not a comment", (string)session.Excel.Run(prefix + "Quote"));
            Assert.AreEqual("hello", (string)session.Excel.Run(prefix + "ThingHello"));
        }

        [Test]
        public void AnExistingFileIsReplacedOrKept()
        {
            var built = BuildWorkbook();
            var output = Path.Combine(_folder, "out");
            Directory.CreateDirectory(output);
            var settings = new ShipSettings { Format = ShipFormat.Xlsm, OutputFolder = output, FileName = "Tool" };
            var service = Service();

            service.Ship(ProjectOf(built).Object, settings, ShipOnExisting.Ask, null);
            var first = HashOf(Path.Combine(output, "Tool.xlsm"));
            Assert.AreEqual("exists", Assert.Throws<ShipRefusedException>(() => service.Ship(ProjectOf(built).Object, settings, ShipOnExisting.Ask, null)).Reason);

            var both = service.Ship(ProjectOf(built).Object, settings, ShipOnExisting.KeepBoth, null);
            Assert.AreEqual(Path.Combine(output, "Tool (2).xlsm"), both.OutputPath);
            Assert.AreEqual(first, HashOf(Path.Combine(output, "Tool.xlsm")), "The first file was not touched.");

            service.Ship(ProjectOf(built).Object, settings, ShipOnExisting.Replace, null);
            //The replaced file stays as the backup.
            CollectionAssert.AreEquivalent(new[] { "Tool.xlsm", "Tool (2).xlsm", "Tool.xlsm.bak" }, Directory.GetFiles(output).Select(Path.GetFileName).ToArray());
            Assert.AreEqual(first, HashOf(Path.Combine(output, "Tool.xlsm.bak")), "The backup is the file that was replaced.");
        }
    }
}
