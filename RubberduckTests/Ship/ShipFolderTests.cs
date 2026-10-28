using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Abstractions;
using Moq;
using NUnit.Framework;
using Rubberduck.Ship;
using Path = System.IO.Path;

namespace RubberduckTests.Ship
{
    /// <summary>Folders are made under a temp root, so the expected paths are written relative to it.</summary>
    [TestFixture]
    [Category("Ship")]
    public class ShipFolderTests
    {
        private class FakeEnvironment : IShipEnvironment
        {
            public Dictionary<string, string> Variables { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> Known { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            public string GetEnvironmentVariable(string name) => Variables.TryGetValue(name, out var value) ? value : null;
            public string KnownFolder(string name) => Known.TryGetValue(name, out var value) ? value : null;
        }

        private string _root;
        private FakeEnvironment _environment;
        private readonly IFileSystem _files = new FileSystem();

        private string P(string relative) => relative.Length == 0 ? _root : Path.Combine(_root, relative.Replace('/', Path.DirectorySeparatorChar));

        private string DocumentFolder => P("Work/Project");

        [SetUp]
        public void SetUp()
        {
            _root = Path.Combine(Path.GetTempPath(), "rd-ship-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_root);
            _environment = new FakeEnvironment();
            _environment.Variables["LOCALAPPDATA"] = P("Users/Ann/AppData/Local");
            _environment.Variables["APPDATA"] = P("Users/Ann/AppData/Roaming");
            _environment.Variables["USERPROFILE"] = P("Users/Ann");
            foreach (var known in ShipFolders.KnownFolders)
            {
                _environment.Known[known] = P("Users/Ann/" + known);
            }
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

        private ShipFolders Folders() => new ShipFolders(_files, _environment);

        private string Resolve(string folder, string documentFolder = "<document>")
            => Folders().Resolve(folder, documentFolder == "<document>" ? DocumentFolder : documentFolder).Path;

        [TestCase("dist", "Work/Project/dist")]
        [TestCase("out/release", "Work/Project/out/release")]
        [TestCase(".", "Work/Project")]
        [TestCase("../dist", "Work/dist")]
        [TestCase("../..", "")]
        [TestCase("../../Shared/Out", "Shared/Out")]
        public void ARelativeFolderIsRelativeToTheFolderOfTheDocument(string written, string expected)
        {
            Assert.AreEqual(P(expected), Resolve(written));
            Assert.AreEqual(P(expected), Resolve(written.Replace('/', '\\')), "Backslashes work too.");
        }

        [Test]
        public void AFullPathIsTakenAsItIs()
        {
            Assert.AreEqual(P("Releases/Tool"), Resolve(P("Releases/Tool")));
            Assert.AreEqual(P("Releases/Tool"), Resolve(P("Releases/Tool").Replace('\\', '/')));
        }

        [Test]
        public void EnvironmentVariablesAreReplaced()
        {
            _environment.Variables["TEMP"] = P("Out");

            Assert.AreEqual(P("Out/Releases"), Resolve("%TEMP%/Releases"));
            Assert.AreEqual(P("Users/Ann/Releases"), Resolve("%USERPROFILE%\\Releases"));
        }

        [Test]
        public void AVariableThatIsNotAFolderIsNeverExpanded_SoASecretCannotBeReadThroughTheFolder()
        {
            _environment.Variables["GITHUB_TOKEN"] = "ghp_secret";

            var result = Folders().Resolve("%GITHUB_TOKEN%", DocumentFolder);

            Assert.IsNull(result.Path);
            StringAssert.DoesNotContain("ghp_secret", result.Problem);
            StringAssert.Contains("GITHUB_TOKEN", result.Problem);
        }

        [Test]
        public void AnUnknownVariableIsNotAFolder()
        {
            var result = Folders().Resolve("%NOPE%/Out", DocumentFolder);

            Assert.IsNull(result.Path);
            StringAssert.Contains("NOPE", result.Problem);
        }

        [TestCase("Documents", "Users/Ann/Documents")]
        [TestCase("Desktop", "Users/Ann/Desktop")]
        [TestCase("Downloads", "Users/Ann/Downloads")]
        [TestCase("downloads", "Users/Ann/Downloads")]
        [TestCase("Music/Albums", "Users/Ann/Music/Albums")]
        [TestCase("Documents\\Releases\\Tool", "Users/Ann/Documents/Releases/Tool")]
        public void AKnownWindowsFolderIsFoundByName(string written, string expected)
        {
            Assert.AreEqual(P(expected), Resolve(written));
        }

        [Test]
        public void ADotSlashMakesAFolderWithTheNameOfAKnownFolderRelative()
        {
            Assert.AreEqual(P("Work/Project/Documents"), Resolve("./Documents"));
        }

        [Test]
        public void OneDriveComesFromTheEnvironment()
        {
            _environment.Variables["OneDrive"] = P("Users/Ann/OneDrive - Contoso");

            Assert.AreEqual(P("Users/Ann/OneDrive - Contoso"), Resolve("OneDrive"));
            Assert.AreEqual(P("Users/Ann/OneDrive - Contoso/Releases"), Resolve("OneDrive/Releases"));
        }

        [Test]
        public void OneDriveFallsBackToTheOtherVariables()
        {
            _environment.Variables["OneDriveConsumer"] = P("Users/Ann/OneDrive");

            Assert.AreEqual(P("Users/Ann/OneDrive"), Resolve("OneDrive"));
        }

        [Test]
        public void WithoutOneDriveTheFolderCannotBeResolvedAndSaysWhy()
        {
            var result = Folders().Resolve("OneDrive/Releases", DocumentFolder);

            Assert.IsNull(result.Path);
            StringAssert.Contains("OneDrive", result.Problem);
        }

        private void DropboxInfo(string variable, string json)
        {
            var folder = Path.Combine(_environment.Variables[variable], "Dropbox");
            Directory.CreateDirectory(folder);
            File.WriteAllText(Path.Combine(folder, "info.json"), json);
        }

        [Test]
        public void DropboxIsTheMainFolderFromItsInfoFile()
        {
            DropboxInfo("LOCALAPPDATA", "{ \"business\": { \"path\": \"D:\\\\Dropbox (Work)\" }, \"personal\": { \"path\": \"D:\\\\Dropbox\" } }");

            Assert.AreEqual("D:\\Dropbox", Resolve("Dropbox"));
            Assert.AreEqual("D:\\Dropbox\\Shipped", Resolve("Dropbox/Shipped"));
        }

        [Test]
        public void DropboxFallsBackToTheRoamingFolderAndToAnAccountThatIsNotPersonal()
        {
            DropboxInfo("APPDATA", "{ \"business\": { \"path\": \"D:\\\\Dropbox (Work)\" } }");

            Assert.AreEqual("D:\\Dropbox (Work)", Resolve("Dropbox"));
        }

        [TestCase(null)]
        [TestCase("not json")]
        [TestCase("{ }")]
        public void WithoutAUsableInfoFileDropboxCannotBeResolved(string json)
        {
            if (json != null)
            {
                DropboxInfo("LOCALAPPDATA", json);
            }

            var result = Folders().Resolve("Dropbox", DocumentFolder);

            Assert.IsNull(result.Path);
            StringAssert.Contains("Dropbox", result.Problem);
        }

        [Test]
        public void NothingSetMeansNoFolder()
        {
            Assert.IsNull(Folders().Resolve(null, DocumentFolder).Path);
            Assert.IsNull(Folders().Resolve("  ", DocumentFolder).Path);
            Assert.IsNull(Folders().Resolve(null, DocumentFolder).Problem);
        }

        [Test]
        public void ARelativeFolderNeedsADocumentFolder()
        {
            var result = Folders().Resolve("dist", null);

            Assert.IsNull(result.Path);
            Assert.IsNotEmpty(result.Problem);
        }

        [Test]
        public void ARealFolderThatCanBeWrittenToPasses()
        {
            Assert.DoesNotThrow(() => Folders().Validate(_root));
            Assert.IsEmpty(Directory.GetFiles(_root), "The probe file must be removed again.");
        }

        [Test]
        public void AFolderThatDoesNotExistIsRefused()
        {
            var missing = P("Nowhere");
            var exception = Assert.Throws<ShipRefusedException>(() => Folders().Validate(missing));

            Assert.AreEqual("folderMissing", exception.Reason);
            StringAssert.Contains(missing, exception.Message);
        }

        [Test]
        public void AFolderThatCannotBeWrittenToIsRefused()
        {
            var file = new Mock<IFile>();
            file.Setup(f => f.Create(It.IsAny<string>())).Throws<UnauthorizedAccessException>();
            var directory = new Mock<IDirectory>();
            directory.Setup(d => d.Exists(It.IsAny<string>())).Returns(true);
            var fileSystem = new Mock<IFileSystem>();
            fileSystem.SetupGet(f => f.File).Returns(file.Object);
            fileSystem.SetupGet(f => f.Directory).Returns(directory.Object);
            fileSystem.SetupGet(f => f.Path).Returns(_files.Path);

            var exception = Assert.Throws<ShipRefusedException>(() => new ShipFolders(fileSystem.Object, _environment).Validate(_root));

            Assert.AreEqual("folderNotWritable", exception.Reason);
        }
    }
}
