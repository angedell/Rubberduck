using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Abstractions;
using System.Linq;
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
    [TestFixture]
    [Category("Ship")]
    public class ShipServiceTests
    {
        private const string CodeWithComments = "Option Explicit\r\n' about Foo\r\nPublic Sub Foo()\r\n    Debug.Print 1 ' one\r\nEnd Sub";
        private const string TestCode = "'@TestModule\r\nPublic Sub A()\r\nEnd Sub";

        private class FakeEnvironment : IShipEnvironment
        {
            public string GetEnvironmentVariable(string name) => Environment.GetEnvironmentVariable(name);
            public string KnownFolder(string name) => null;
        }

        /// <summary>Stands in for Excel: records what it was asked to build and writes a file with known content.</summary>
        private class FakeWriter : IShipArtifactWriter
        {
            public List<ShipArtifactJob> Jobs { get; } = new List<ShipArtifactJob>();
            public Exception Failure { get; set; }
            public string Content { get; set; } = "NEW ARTIFACT";

            public void Write(ShipArtifactJob job, IProgress<string> progress)
            {
                Jobs.Add(job);
                File.WriteAllText(job.ArtifactPath, Content);
                if (Failure != null)
                {
                    throw Failure;
                }
            }
        }

        private string _root;
        private string _workbookFolder;
        private string _outputFolder;
        private string _workbook;
        private FakeWriter _writer;
        private Mock<IPackageResourceService> _resources;
        private Mock<IVBProject> _project;
        private List<Mock<IVBComponent>> _components;
        private List<Mock<ICodeModule>> _codeModules;

        [SetUp]
        public void SetUp()
        {
            _root = Path.Combine(Path.GetTempPath(), "rd-shipsvc-" + Guid.NewGuid().ToString("N"));
            _workbookFolder = Path.Combine(_root, "work");
            _outputFolder = Path.Combine(_root, "out");
            Directory.CreateDirectory(_workbookFolder);
            Directory.CreateDirectory(_outputFolder);
            _workbook = Path.Combine(_workbookFolder, "Book.xlsm");
            File.WriteAllText(_workbook, "SAVED WORKBOOK");
            _writer = new FakeWriter();
            _resources = new Mock<IPackageResourceService>();
            BuildProject(isSaved: true, host: "Excel", fileName: _workbook);
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

        private void BuildProject(bool isSaved, string host, string fileName)
        {
            var builder = new MockVbeBuilder().ProjectBuilder("MyBook", fileName, ProjectProtection.Unprotected)
                .AddComponent("ModCode", ComponentType.StandardModule, CodeWithComments)
                .AddComponent("TestMath", ComponentType.StandardModule, TestCode)
                .AddComponent("Note_1", ComponentType.StandardModule, "'@Notes(\"README.md\")\r\n' hello")
                .AddComponent("ThisWorkbook", ComponentType.Document, "' workbook\r\nOption Explicit");
            _components = builder.MockComponents;
            _codeModules = builder.MockCodeModules;
            _project = builder.Build();
            _project.SetupGet(p => p.IsSaved).Returns(isSaved);
            _project.SetupGet(p => p.ProjectId).Returns("project-1");

            var hostApplication = new Mock<IHostApplication>();
            hostApplication.SetupGet(h => h.ApplicationName).Returns(host);
            var vbe = Mock.Get(_project.Object.VBE);
            vbe.Setup(v => v.HostApplication()).Returns(hostApplication.Object);

            //What the VBE writes when a component is exported: the attribute line, then the code.
            foreach (var component in _components)
            {
                var mock = component;
                mock.Setup(c => c.Export(It.IsAny<string>())).Callback<string>(path =>
                    File.WriteAllText(path, "Attribute VB_Name = \"" + mock.Object.Name + "\"\r\n" + mock.Object.CodeModule.Content(), Encoding.Default));
            }
        }

        private ShipService Service(IShipEnvironment environment = null, IShipAnalysis analysis = null)
            => new ShipService(new FileSystem(), _resources.Object, ShipTestParser.Create(), _writer, environment ?? new FakeEnvironment(),
                Path.Combine(_root, "settings", "ship-unattended.json"), Path.Combine(_root, "temp"), analysis);

        private class FakeAnalysis : IShipAnalysis
        {
            public IVBProject AskedAbout { get; private set; }
            public ShipWorkingSet Seen { get; private set; }
            public ShipAnalysisResult Result { get; set; }
            public string WhyNot { get; set; }

            public ShipAnalysisResult Analyze(IVBProject project, ShipWorkingSet set, ShipTreeShakingSettings settings, out string whyNot)
            {
                AskedAbout = project;
                Seen = set;
                whyNot = WhyNot;
                return Result;
            }
        }

        private ShipSettings Settings(ShipOnExisting onExisting = ShipOnExisting.Ask, string folder = null)
            => new ShipSettings { OutputFolder = folder ?? _outputFolder, OnExisting = onExisting, FileName = "Tool" };

        private static string Hash(string text)
        {
            using (var sha = SHA256.Create())
            {
                return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(text)));
            }
        }

        #region Rules

        [Test]
        public void ExcelProjectWithAFileCanBeShipped()
        {
            Assert.IsTrue(Service().CanShip(_project.Object));
        }

        [Test]
        public void AnotherHostCannotBeShipped()
        {
            BuildProject(true, "Word", Path.Combine(_workbookFolder, "Doc.docm"));

            Assert.IsFalse(Service().CanShip(_project.Object));
            Assert.AreEqual("notExcel", Assert.Throws<ShipRefusedException>(() => Service().Ship(_project.Object, Settings(), ShipOnExisting.Ask, null)).Reason);
        }

        [Test]
        public void AProjectWithoutAFileCannotBeShipped()
        {
            BuildProject(false, "Excel", string.Empty);

            Assert.IsFalse(Service().CanShip(_project.Object));
            Assert.AreEqual("unsaved", Assert.Throws<ShipRefusedException>(() => Service().Ship(_project.Object, Settings(), ShipOnExisting.Ask, null)).Reason);
        }

        [Test]
        public void AProjectWithUnsavedChangesIsRefusedBecauseTheCopyIsOfTheFile()
        {
            BuildProject(false, "Excel", _workbook);

            var exception = Assert.Throws<ShipRefusedException>(() => Service().Ship(_project.Object, Settings(), ShipOnExisting.Ask, null));

            Assert.AreEqual("unsaved", exception.Reason);
            Assert.IsEmpty(_writer.Jobs);
        }

        [Test]
        public void PreviewIsRefusedForTheSameReasons()
        {
            BuildProject(false, "Excel", _workbook);

            Assert.AreEqual("unsaved", Assert.Throws<ShipRefusedException>(() => Service().Preview(_project.Object, Settings())).Reason);
        }

        [Test]
        public void AFolderThatIsNotThereIsRefused()
        {
            var exception = Assert.Throws<ShipRefusedException>(() => Service().Ship(_project.Object, Settings(folder: Path.Combine(_root, "nowhere")), ShipOnExisting.Ask, null));

            Assert.AreEqual("folderMissing", exception.Reason);
        }

        [Test]
        public void NoFolderAtAllIsRefusedWhenShippingButNotWhenPreviewing()
        {
            var settings = new ShipSettings();

            Assert.AreEqual("folderMissing", Assert.Throws<ShipRefusedException>(() => Service().Ship(_project.Object, settings, ShipOnExisting.Ask, null)).Reason);
            Assert.DoesNotThrow(() => Service().Preview(_project.Object, settings));
        }

        [Test]
        public void AFolderThatIsSetButBadIsRefusedByPreviewToo()
        {
            Assert.AreEqual("folderMissing",
                Assert.Throws<ShipRefusedException>(() => Service().Preview(_project.Object, Settings(folder: Path.Combine(_root, "nowhere")))).Reason);
        }

        [Test]
        public void TheFolderIsResolvedRelativeToTheWorkbook()
        {
            Assert.AreEqual(_outputFolder, Service().ResolveOutputFolder(_project.Object, new ShipSettings { OutputFolder = "../out" }));
            Assert.IsNull(Service().ResolveOutputFolder(_project.Object, new ShipSettings()));
        }

        #endregion

        #region The artifact

        [Test]
        public void ShippingBuildsTheArtifactFromTheStrippedProjectAndWritesItToTheFolder()
        {
            var report = Service().Ship(_project.Object, Settings(), ShipOnExisting.Ask, null);

            var path = Path.Combine(_outputFolder, "Tool.xlsm");
            Assert.AreEqual(path, report.OutputPath);
            Assert.AreEqual("NEW ARTIFACT", File.ReadAllText(path));
            Assert.IsFalse(report.DryRun);

            var job = _writer.Jobs.Single();
            Assert.AreEqual(_workbook, job.SourceFile);
            Assert.AreEqual(ShipFormat.Xlsm, job.Format);
            Assert.AreEqual("Tool.xlsm", Path.GetFileName(job.ArtifactPath));
            Assert.IsTrue(job.Set.Modules.Single(m => m.Name == "TestMath").Removed);
            Assert.IsTrue(job.Set.Modules.Single(m => m.Name == "Note_1").Removed);
            Assert.AreEqual("Attribute VB_Name = \"ModCode\"\r\nOption Explicit\r\nPublic Sub Foo()\r\n    Debug.Print 1\r\nEnd Sub", job.Set.Modules.Single(m => m.Name == "ModCode").Text);
            Assert.AreEqual(ShipSteps.All.Where(step => !ShipSteps.OffByDefault(step)).ToArray(), report.Steps.Select(s => s.Step).ToArray(), "Tree shaking is off unless asked for.");
        }

        [Test]
        public void TheFormatDecidesTheExtension()
        {
            var settings = Settings();
            settings.Format = ShipFormat.Xlam;

            var report = Service().Ship(_project.Object, settings, ShipOnExisting.Ask, null);

            Assert.AreEqual(Path.Combine(_outputFolder, "Tool.xlam"), report.OutputPath);
            Assert.AreEqual(ShipFormat.Xlam, _writer.Jobs.Single().Format);
        }

        [Test]
        public void WithoutAFileNameTheProjectNameIsUsed()
        {
            var settings = Settings();
            settings.FileName = null;

            Assert.AreEqual(Path.Combine(_outputFolder, "MyBook.xlsm"), Service().Ship(_project.Object, settings, ShipOnExisting.Ask, null).OutputPath);
        }

        [Test]
        public void CharactersThatCannotBeInAFileNameAreReplaced()
        {
            var settings = Settings();
            settings.FileName = "My:Tool?";

            Assert.AreEqual("My_Tool_.xlsm", Path.GetFileName(Service().Ship(_project.Object, settings, ShipOnExisting.Ask, null).OutputPath));
        }

        [Test]
        public void TheStepsThatAreSwitchedOffAreNotRun()
        {
            var settings = Settings();
            settings.Steps[ShipSteps.Comments] = false;

            var report = Service().Ship(_project.Object, settings, ShipOnExisting.Ask, null);

            CollectionAssert.DoesNotContain(report.Steps.Select(s => s.Step), ShipSteps.Comments);
            StringAssert.Contains("' about Foo", _writer.Jobs.Single().Set.Modules.Single(m => m.Name == "ModCode").Text);
        }

        [Test]
        public void ThePreviewReportsWhatWouldBeRemovedAndWritesNothing()
        {
            var before = Directory.GetFiles(_root, "*", SearchOption.AllDirectories).OrderBy(x => x).ToList();

            var report = Service().Preview(_project.Object, Settings());

            Assert.IsTrue(report.DryRun);
            Assert.IsEmpty(_writer.Jobs, "Preview does not build anything.");
            Assert.AreEqual(Path.Combine(_outputFolder, "Tool.xlsm"), report.OutputPath);
            Assert.AreEqual(3, report.Steps.Single(s => s.Step == ShipSteps.Comments).Removed, "Two comments in ModCode and one in ThisWorkbook; the removed modules are not counted.");
            Assert.AreEqual(1, report.Steps.Single(s => s.Step == ShipSteps.Notes).Removed);
            Assert.AreEqual(1, report.Steps.Single(s => s.Step == ShipSteps.TestModules).Removed);
            CollectionAssert.AreEqual(before, Directory.GetFiles(_root, "*", SearchOption.AllDirectories).OrderBy(x => x).ToList(), "No file is left anywhere, the temporary ones included.");
        }

        [Test]
        public void TheTemporaryFilesAreGoneAfterwardsEvenWhenTheBuildFails()
        {
            _writer.Failure = new InvalidOperationException("Excel died");

            Assert.Throws<InvalidOperationException>(() => Service().Ship(_project.Object, Settings(), ShipOnExisting.Ask, null));

            Assert.IsEmpty(Directory.GetFiles(_outputFolder), "Nothing half-written in the output folder.");
            var temp = Path.Combine(_root, "temp");
            Assert.That(!Directory.Exists(temp) || !Directory.GetFileSystemEntries(temp).Any(), "The work folder must be deleted.");
        }

        [Test]
        public void ARefusalFromTheWriterIsPassedOn()
        {
            _writer.Failure = new ShipRefusedException("vbomNotTrusted", "Trust access to the VBA project object model.");

            Assert.AreEqual("vbomNotTrusted", Assert.Throws<ShipRefusedException>(() => Service().Ship(_project.Object, Settings(), ShipOnExisting.Ask, null)).Reason);
        }

        [Test]
        public void ProgressIsReported()
        {
            var messages = new List<string>();

            Service().Ship(_project.Object, Settings(), ShipOnExisting.Ask, new Progress<string>(messages.Add));
            System.Threading.Thread.Sleep(100);

            Assert.IsNotEmpty(messages);
        }

        #endregion

        #region A file that is already there

        private string Existing()
        {
            var path = Path.Combine(_outputFolder, "Tool.xlsm");
            File.WriteAllText(path, "OLD ARTIFACT");
            return path;
        }

        [Test]
        public void AskRefusesAndNamesTheFileInTheWay()
        {
            var existing = Existing();

            var exception = Assert.Throws<ShipRefusedException>(() => Service().Ship(_project.Object, Settings(), ShipOnExisting.Ask, null));

            Assert.AreEqual("exists", exception.Reason);
            Assert.AreEqual(existing, exception.Path);
            Assert.AreEqual("OLD ARTIFACT", File.ReadAllText(existing));
            Assert.IsEmpty(_writer.Jobs, "Nothing is built before the answer is known.");
        }

        [Test]
        public void ReplaceSwapsTheFileAndKeepsTheOldOneAsTheBackup()
        {
            var existing = Existing();

            var report = Service().Ship(_project.Object, Settings(), ShipOnExisting.Replace, null);

            Assert.AreEqual(existing, report.OutputPath);
            Assert.AreEqual("NEW ARTIFACT", File.ReadAllText(existing));
            Assert.AreEqual("OLD ARTIFACT", File.ReadAllText(existing + ".bak"));
            Assert.AreEqual(2, Directory.GetFiles(_outputFolder).Length);
        }

        [Test]
        public void ReplaceNeverOverwritesABackupThatIsAlreadyThere()
        {
            var existing = Existing();
            File.WriteAllText(existing + ".bak", "MADE BY HAND");

            Service().Ship(_project.Object, Settings(), ShipOnExisting.Replace, null);

            Assert.AreEqual("MADE BY HAND", File.ReadAllText(existing + ".bak"));
            Assert.AreEqual("OLD ARTIFACT", File.ReadAllText(existing + ".bak2"));
        }

        [Test]
        public void TheWorkbookItselfIsNeverTheTarget()
        {
            var settings = Settings();
            settings.OutputFolder = Path.GetDirectoryName(_project.Object.FileName);
            settings.FileName = Path.GetFileNameWithoutExtension(_project.Object.FileName);

            var exception = Assert.Throws<ShipRefusedException>(() => Service().Ship(_project.Object, settings, ShipOnExisting.Replace, null));

            Assert.AreEqual("targetIsSource", exception.Reason);
        }

        [Test]
        public void AFolderThatLoadsFilesByItselfIsRefused()
        {
            var startup = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Excel", "XLSTART");
            var settings = Settings();
            settings.OutputFolder = startup;

            var exception = Assert.Throws<ShipRefusedException>(() => Service().Ship(_project.Object, settings, ShipOnExisting.KeepBoth, null));

            Assert.AreEqual("restrictedFolder", exception.Reason);
        }

        [Test]
        public void ReplaceKeepsTheOldFileIfTheNewOneCannotBeBuilt()
        {
            var existing = Existing();
            _writer.Failure = new InvalidOperationException("Excel died");

            Assert.Throws<InvalidOperationException>(() => Service().Ship(_project.Object, Settings(), ShipOnExisting.Replace, null));

            Assert.AreEqual("OLD ARTIFACT", File.ReadAllText(existing));
            CollectionAssert.AreEqual(new[] { existing }, Directory.GetFiles(_outputFolder));
        }

        [Test]
        public void ReplaceKeepsTheOldFileIfItIsLockedAndSaysSo()
        {
            var existing = Existing();

            using (new FileStream(existing, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                var exception = Assert.Catch<IOException>(() => Service().Ship(_project.Object, Settings(), ShipOnExisting.Replace, null));
                StringAssert.Contains("Tool.xlsm", exception.Message);
            }

            Assert.AreEqual("OLD ARTIFACT", File.ReadAllText(existing));
            CollectionAssert.AreEqual(new[] { existing }, Directory.GetFiles(_outputFolder), "No temporary or backup file is left behind.");
        }

        [Test]
        public void KeepBothNumbersTheNewFile()
        {
            var existing = Existing();

            var first = Service().Ship(_project.Object, Settings(), ShipOnExisting.KeepBoth, null);
            var second = Service().Ship(_project.Object, Settings(), ShipOnExisting.KeepBoth, null);

            Assert.AreEqual(Path.Combine(_outputFolder, "Tool (2).xlsm"), first.OutputPath);
            Assert.AreEqual(Path.Combine(_outputFolder, "Tool (3).xlsm"), second.OutputPath);
            Assert.AreEqual("OLD ARTIFACT", File.ReadAllText(existing));
            Assert.AreEqual("Tool (3).xlsm", Path.GetFileName(_writer.Jobs.Last().ArtifactPath), "The add-in is built under the name it will have.");
        }

        [Test]
        public void TheCallersAnswerBeatsTheSettingAndAskInTheCallDefersToTheSetting()
        {
            Existing();

            Assert.DoesNotThrow(() => Service().Ship(_project.Object, Settings(ShipOnExisting.Ask), ShipOnExisting.Replace, null));
            Assert.DoesNotThrow(() => Service().Ship(_project.Object, Settings(ShipOnExisting.KeepBoth), ShipOnExisting.Ask, null));
            Assert.AreEqual("exists", Assert.Throws<ShipRefusedException>(() => Service().Ship(_project.Object, Settings(ShipOnExisting.Ask), ShipOnExisting.Ask, null)).Reason);
        }

        #endregion

        #region The working project

        [Test]
        public void TheWorkingProjectIsUntouchedByAShip()
        {
            string Snapshot() => string.Join("|", _components.Select(c => c.Object.Name + ":" + Hash(c.Object.CodeModule.Content())));
            var before = Snapshot();

            Service().Ship(_project.Object, Settings(), ShipOnExisting.Ask, null);
            Service().Preview(_project.Object, Settings());

            Assert.AreEqual(before, Snapshot(), "The text of every module is the same.");
            Assert.AreEqual(4, _components.Count);
            foreach (var codeModule in _codeModules)
            {
                codeModule.Verify(m => m.AddFromString(It.IsAny<string>()), Times.Never);
                codeModule.Verify(m => m.AddFromFile(It.IsAny<string>()), Times.Never);
                codeModule.Verify(m => m.InsertLines(It.IsAny<int>(), It.IsAny<string>()), Times.Never);
                codeModule.Verify(m => m.ReplaceLine(It.IsAny<int>(), It.IsAny<string>()), Times.Never);
                codeModule.Verify(m => m.DeleteLines(It.IsAny<int>(), It.IsAny<int>()), Times.Never);
                codeModule.Verify(m => m.Clear(), Times.Never);
            }

            Mock.Get(_project.Object.VBComponents).Verify(c => c.Remove(It.IsAny<IVBComponent>()), Times.Never);
            Mock.Get(_project.Object.VBComponents).Verify(c => c.Import(It.IsAny<string>()), Times.Never);
            Mock.Get(_project.Object.VBComponents).Verify(c => c.Add(It.IsAny<ComponentType>()), Times.Never);
            Mock.Get(_project.Object.VBComponents).Verify(c => c.RemoveSafely(It.IsAny<IVBComponent>()), Times.Never);
            Mock.Get(_project.Object.References).Verify(r => r.Remove(It.IsAny<IReference>()), Times.Never);
            _project.Verify(p => p.SaveAs(It.IsAny<string>()), Times.Never);
            Assert.AreEqual("SAVED WORKBOOK", File.ReadAllText(_workbook), "The saved file is not touched either.");
        }

        #endregion

        #region Settings

        private const string ConfigWithShip = "{\r\n  // team\r\n  \"ship\": { \"format\": \"xlam\", \"outputFolder\": \"../out\", \"fileName\": \"Tool\" }\r\n}";

        private void ConfigIs(string text)
        {
            _resources.Setup(r => r.IsSupported(_workbook)).Returns(true);
            if (text == null)
            {
                _resources.Setup(r => r.Read(_workbook, "rubberduck.json")).Throws(new PackageResourceException(PackageResourceError.NotFound, "none"));
            }
            else
            {
                _resources.Setup(r => r.Read(_workbook, "rubberduck.json")).Returns(new PackageResourceContent { Name = "rubberduck.json", Text = text, Hash = "h1" });
            }
        }

        [Test]
        public void SettingsAreReadFromTheDocumentsConfig()
        {
            ConfigIs(ConfigWithShip);

            var settings = Service().ReadSettings(_project.Object);

            Assert.AreEqual(ShipFormat.Xlam, settings.Format);
            Assert.AreEqual("../out", settings.OutputFolder);
            Assert.AreEqual("Tool", settings.FileName);
        }

        [Test]
        public void WithoutAConfigOrAShipSectionTheDefaultsApply()
        {
            ConfigIs(null);
            Assert.AreEqual(ShipFormat.Xlsm, Service().ReadSettings(_project.Object).Format);

            ConfigIs("{ \"sync\": {} }");
            Assert.IsNull(Service().ReadSettings(_project.Object).OutputFolder);
        }

        [Test]
        public void WritingSettingsStagesTheConfigWithOnlyTheShipSectionChanged()
        {
            ConfigIs("{\r\n  // team\r\n  \"sync\": { \"onViolation\": \"warn\" }\r\n}");
            string staged = null;
            _resources.Setup(r => r.SetConfig(_workbook, It.IsAny<string>(), "h1")).Callback<string, string, string>((path, json, hash) => staged = json).Returns(new PackageResource());

            Service().WriteSettings(_project.Object, new ShipSettings { OutputFolder = "dist", Format = ShipFormat.Xlam });

            StringAssert.Contains("// team", staged);
            StringAssert.Contains("\"onViolation\": \"warn\"", staged);
            var read = new Rubberduck.Settings.ProjectConfigReader(null).Parse(staged, null);
            Assert.IsEmpty(read.Errors);
            Assert.AreEqual("dist", read.Ship.OutputFolder);
            Assert.AreEqual(ShipFormat.Xlam, read.Ship.Format);
        }

        [Test]
        public void WritingSettingsToADocumentWithoutAConfigCreatesOne()
        {
            ConfigIs(null);
            string staged = null;
            _resources.Setup(r => r.SetConfig(_workbook, It.IsAny<string>(), null)).Callback<string, string, string>((path, json, hash) => staged = json).Returns(new PackageResource());

            Service().WriteSettings(_project.Object, new ShipSettings { OutputFolder = "dist" });

            var read = new Rubberduck.Settings.ProjectConfigReader(null).Parse(staged, null);
            Assert.IsEmpty(read.Errors);
            Assert.AreEqual("dist", read.Ship.OutputFolder);
        }

        #endregion

        #region Unattended ship

        [Test]
        public void UnattendedShipIsOffUntilTheUserSwitchesItOn()
        {
            Assert.IsFalse(Service().IsUnattended(_project.Object));
        }

        [Test]
        public void TheSwitchIsKeptPerDocumentInTheUsersOwnFileAndNotInTheConfig()
        {
            Service().SetUnattended(_project.Object, true);

            Assert.IsTrue(Service().IsUnattended(_project.Object), "A new service reads it from the file.");
            var file = Path.Combine(_root, "settings", "ship-unattended.json");
            StringAssert.Contains("Book.xlsm", File.ReadAllText(file));
            _resources.Verify(r => r.SetConfig(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>()), Times.Never);

            BuildProject(true, "Excel", Path.Combine(_workbookFolder, "Other.xlsm"));
            Assert.IsFalse(Service().IsUnattended(_project.Object), "Another document is not affected.");

            BuildProject(true, "Excel", _workbook);
            Service().SetUnattended(_project.Object, false);
            Assert.IsFalse(Service().IsUnattended(_project.Object));
        }

        [Test]
        public void TheSwitchCannotBeSetThroughTheConfig()
        {
            ConfigIs("{ \"ship\": { \"unattended\": true } }");

            Assert.IsFalse(Service().IsUnattended(_project.Object));
            Assert.That(new Rubberduck.Settings.ProjectConfigReader(null).Parse("{ \"ship\": { \"unattended\": true } }", null).Errors.Any(e => e.Contains("unattended")),
                "An agent that writes it into the file gets an error, not a switch.");
        }

        #endregion

        #region Tree shaking

        [Test]
        public void TreeShakingSwitchedOnWithoutAnAnalysisIsAWarningNotACrash()
        {
            var settings = Settings();
            settings.Steps[ShipSteps.TreeShaking] = true;

            var report = Service().Preview(_project.Object, settings);

            Assert.That(report.Warnings, Has.Some.StartsWith("Tree shaking was skipped"));
            Assert.IsNull(report.Steps.SingleOrDefault(s => s.Step == ShipSteps.TreeShaking));
            Assert.IsNotNull(report.Steps.SingleOrDefault(s => s.Step == ShipSteps.Comments), "The other steps still run.");
        }

        [Test]
        public void ThePreviewAsksTheAnalysisAboutTheProjectAndCarriesItsRootsAndDynamicCalls()
        {
            var analysis = new FakeAnalysis { Result = new ShipAnalysisResult { Roots = { "ModCode.Foo: test" }, DynamicCalls = { "ModCode.Foo: Application.Run" } } };
            var settings = Settings();
            settings.Steps[ShipSteps.TreeShaking] = true;

            var report = Service(analysis: analysis).Preview(_project.Object, settings);

            Assert.AreSame(_project.Object, analysis.AskedAbout);
            CollectionAssert.Contains(analysis.Seen.Modules.Select(m => m.Name), "ModCode");
            CollectionAssert.AreEqual(new[] { "ModCode.Foo: test" }, report.Roots);
            CollectionAssert.AreEqual(new[] { "ModCode.Foo: Application.Run" }, report.DynamicCalls);
            Assert.IsNotNull(report.Steps.SingleOrDefault(s => s.Step == ShipSteps.TreeShaking));
        }

        [Test]
        public void AnAnalysisThatSaysNoLeavesTheShipGoingWithTheReasonAsAWarning()
        {
            var analysis = new FakeAnalysis { WhyNot = "module ModCode does not parse" };
            var settings = Settings();
            settings.Steps[ShipSteps.TreeShaking] = true;

            var report = Service(analysis: analysis).Ship(_project.Object, settings, ShipOnExisting.Ask, null);

            Assert.That(report.Warnings, Has.Some.Contain("module ModCode does not parse"));
            Assert.IsNotNull(report.OutputPath);
            Assert.AreEqual(1, _writer.Jobs.Count, "The artifact is still written.");
        }

        [Test]
        public void TreeShakingOffNeverAsksTheAnalysis()
        {
            var analysis = new FakeAnalysis { Result = new ShipAnalysisResult() };

            Service(analysis: analysis).Preview(_project.Object, Settings());

            Assert.IsNull(analysis.AskedAbout);
        }

        #endregion
    }
}
