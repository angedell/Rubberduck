using System.IO;
using System.IO.Abstractions;
using System.Linq;
using NUnit.Framework;
using Rubberduck.UI.CodeExplorer.GitSync;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    [Category("GitSync")]
    public class GitSyncTextTests
    {
        [Test]
        [TestCase("a\nb", "a\r\nb")]
        [TestCase("a\r\nb", "a\r\nb")]
        [TestCase("a\rb", "a\r\nb")]
        [TestCase("a\r\nb\nc\r\n", "a\r\nb\r\nc\r\n")]
        public void NormalizeLineEndings_ConvertsToCrlf(string input, string expected)
        {
            Assert.AreEqual(expected, GitSyncText.NormalizeLineEndings(input));
        }

        [Test]
        [TestCase("", "")]
        [TestCase("Models", "Models")]
        [TestCase("Models\\Orders", "Models.Orders")]
        [TestCase("Models/Orders/Lines", "Models.Orders.Lines")]
        [TestCase("v1.2\\Core", "v1_2.Core")]
        public void FolderPath_FromRelativeDirectory(string directory, string expected)
        {
            Assert.AreEqual(expected, GitSyncText.FolderPath(directory));
        }

        [Test]
        [TestCase("", "")]
        [TestCase("Models.Orders", "Models\\Orders")]
        public void RelativeDirectory_FromFolderPath(string folderPath, string expected)
        {
            Assert.AreEqual(expected, GitSyncText.RelativeDirectory(folderPath));
        }

        [Test]
        [TestCase("'@Folder(\"A.B\")\r\nOption Explicit", "A.B")]
        [TestCase("'@Folder \"A.B\"\r\nOption Explicit", "A.B")]
        [TestCase("'@folder(\"A.B\")", "A.B")]
        [TestCase("Option Explicit", null)]
        public void GetFolderAnnotation(string source, string expected)
        {
            Assert.AreEqual(expected, GitSyncText.GetFolderAnnotation(source));
        }

        [Test]
        public void SetFolderAnnotation_ReplacesExisting()
        {
            var result = GitSyncText.SetFolderAnnotation("'@Folder(\"Old\")\r\nOption Explicit", "New.Path");
            Assert.AreEqual("'@Folder(\"New.Path\")\r\nOption Explicit", result);
        }

        [Test]
        public void SetFolderAnnotation_LeavesCorrectAnnotationAlone()
        {
            var source = "'@Folder(\"A\")\r\nOption Explicit";
            Assert.AreEqual(source, GitSyncText.SetFolderAnnotation(source, "A"));
        }

        [Test]
        public void SetFolderAnnotation_InsertsAfterAttributes()
        {
            var source = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1\r\nEND\r\nAttribute VB_Name = \"Foo\"\r\nAttribute VB_Exposed = False\r\nOption Explicit";
            var expected = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1\r\nEND\r\nAttribute VB_Name = \"Foo\"\r\nAttribute VB_Exposed = False\r\n'@Folder(\"A.B\")\r\nOption Explicit";
            Assert.AreEqual(expected, GitSyncText.SetFolderAnnotation(source, "A.B"));
        }

        [Test]
        public void SetFolderAnnotation_InsertsAtTopWithoutAttributes()
        {
            Assert.AreEqual("'@Folder(\"A\")\r\nOption Explicit", GitSyncText.SetFolderAnnotation("Option Explicit", "A"));
        }
    }

    [TestFixture]
    [Category("GitSync")]
    public class GitSyncFolderProcessorTests
    {
        private string _root;
        private readonly IFileSystem _fileSystem = new FileSystem();

        [SetUp]
        public void SetUp()
        {
            _root = Path.Combine(Path.GetTempPath(), "RdGitSyncTests", Path.GetRandomFileName());
            Directory.CreateDirectory(_root);
        }

        [TearDown]
        public void TearDown()
        {
            Directory.Delete(_root, true);
        }

        private string Write(string relativePath, string content)
        {
            var path = Path.Combine(_root, relativePath);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, content, GitSyncFolderProcessor.SourceEncoding);
            return path;
        }

        [Test]
        public void PrepareAll_FixesLineEndingsAndAnnotationInSubfolders()
        {
            var path = Write("Models\\Orders\\Foo.cls", "Attribute VB_Name = \"Foo\"\nOption Explicit\n");

            var prepared = new GitSyncFolderProcessor(_fileSystem).PrepareAll(_root).Single();

            Assert.IsTrue(prepared.LineEndingsFixed);
            Assert.IsTrue(prepared.FolderAnnotationFixed);
            Assert.AreEqual("Models\\Orders", prepared.RelativeDirectory);
            Assert.AreEqual("Attribute VB_Name = \"Foo\"\r\n'@Folder(\"Models.Orders\")\r\nOption Explicit\r\n",
                File.ReadAllText(path, GitSyncFolderProcessor.SourceEncoding));
        }

        [Test]
        public void Prepare_CorrectsWrongFolderAnnotation()
        {
            var path = Write("Sub\\Foo.bas", "Attribute VB_Name = \"Foo\"\r\n'@Folder(\"Elsewhere\")\r\nOption Explicit\r\n");

            new GitSyncFolderProcessor(_fileSystem).Prepare(_root, path);

            StringAssert.Contains("'@Folder(\"Sub\")", File.ReadAllText(path, GitSyncFolderProcessor.SourceEncoding));
            StringAssert.DoesNotContain("Elsewhere", File.ReadAllText(path, GitSyncFolderProcessor.SourceEncoding));
        }

        [Test]
        public void Prepare_LeavesRootFilesAnnotationAlone()
        {
            var content = "Attribute VB_Name = \"Foo\"\r\nOption Explicit\r\n";
            var path = Write("Foo.bas", content);

            var prepared = new GitSyncFolderProcessor(_fileSystem).Prepare(_root, path);

            Assert.IsFalse(prepared.LineEndingsFixed);
            Assert.IsFalse(prepared.FolderAnnotationFixed);
            Assert.AreEqual(content, File.ReadAllText(path, GitSyncFolderProcessor.SourceEncoding));
        }

        [Test]
        public void SourceFiles_SkipsHiddenFoldersAndOtherFiles()
        {
            Write("A\\Foo.cls", "x");
            Write(".git\\Bar.cls", "x");
            Write("A\\readme.md", "x");

            var files = new GitSyncFolderProcessor(_fileSystem).SourceFiles(_root).ToList();

            Assert.AreEqual(1, files.Count);
            StringAssert.EndsWith("Foo.cls", files[0]);
        }
    }
}
