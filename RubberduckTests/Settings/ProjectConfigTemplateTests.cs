using System;
using System.IO;
using System.IO.Abstractions;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.CodeExplorer.GitSync;

namespace RubberduckTests.Settings
{
    [TestFixture]
    public class ProjectConfigTemplateTests
    {
        private string _root;

        [SetUp]
        public void SetUp()
        {
            _root = Path.Combine(Path.GetTempPath(), "rdtemplate-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_root);
        }

        [TearDown]
        public void TearDown() => Directory.Delete(_root, true);

        private static IndenterSettings Unusual() => new IndenterSettings(true)
        {
            IndentSpaces = 3,
            IndentCase = true,
            AlignDims = true,
            AlignDimColumn = 22,
            ForceDebugPrintInColumn1 = true,
            EndOfLineCommentStyle = EndOfLineCommentStyle.SameGap,
            EmptyLineHandlingMethod = EmptyLineHandling.Remove,
            LinesBetweenProcedures = 2,
            GroupRelatedProperties = true,
            MaximumConsecutiveEmptyLines = 2,
            RemoveEmptyLinesAtProcedureEdges = true,
            RemoveLeadingEmptyLines = true
        };

        [Test]
        [Category("Settings")]
        public void TemplateMentionsEveryIndenterSetting()
        {
            CollectionAssert.IsEmpty(ProjectConfigTemplate.Missing(), "a new indenter setting needs an entry and a comment in the template");
        }

        [Test]
        [Category("Settings")]
        public void TemplateReadsBackAsTheSettingsItWasMadeFrom()
        {
            var user = Unusual();
            File.WriteAllText(Path.Combine(_root, ProjectConfig.FileName), ProjectConfigTemplate.Create(user));

            var config = new ProjectConfigReader(new FileSystem()).Read(_root, new IndenterSettings(true));

            CollectionAssert.IsEmpty(config.Errors);
            Assert.AreEqual(SignatureCache.KeyOf(user), SignatureCache.KeyOf(config.Indenter));
        }

        [Test]
        [Category("Settings")]
        public void TemplateHasCommentsAndDefaultSyncSettings()
        {
            var text = ProjectConfigTemplate.Create(new IndenterSettings(true));
            File.WriteAllText(Path.Combine(_root, ProjectConfig.FileName), text);

            var config = new ProjectConfigReader(new FileSystem()).Read(_root, new IndenterSettings(true));

            Assert.GreaterOrEqual(text.Split('\n').Count(line => line.TrimStart().StartsWith("//")), 40);
            CollectionAssert.IsEmpty(config.Sync.Ignore);
            CollectionAssert.IsEmpty(config.Sync.Enforce);
            Assert.AreEqual(ViolationMode.Fix, config.Sync.OnViolation);
        }

        [Test]
        [Category("Settings")]
        public void CreateDefaultWritesOnceAndNeverOverwrites()
        {
            var writer = new ProjectConfigWriter(new FileSystem());

            var first = writer.CreateDefault(_root, Unusual());
            File.AppendAllText(first.Path, "// mine\r\n");
            var second = writer.CreateDefault(_root, new IndenterSettings(true));

            Assert.IsTrue(first.Created);
            Assert.IsFalse(second.Created);
            StringAssert.Contains("// mine", File.ReadAllText(second.Path));
        }

        [Test]
        [Category("Settings")]
        public void CreateDefaultMakesTheFolderIfNeeded()
        {
            var folder = Path.Combine(_root, "new", "deeper");

            var created = new ProjectConfigWriter(new FileSystem()).CreateDefault(folder, new IndenterSettings(true));

            Assert.IsTrue(File.Exists(created.Path));
        }

        [Test]
        [Category("Settings")]
        public void AddingIgnoredKeepsCommentsAndOtherSettings()
        {
            var path = Path.Combine(_root, ProjectConfig.FileName);
            var writer = new ProjectConfigWriter(new FileSystem());
            writer.CreateDefault(_root, Unusual());

            writer.AddIgnored(path, _root, new[] { "/a.bas", "Sandbox/" });
            writer.AddIgnored(path, _root, new[] { "/b.cls", "/a.bas" });

            var text = File.ReadAllText(path);
            var config = new ProjectConfigReader(new FileSystem()).Read(_root, new IndenterSettings(true));
            CollectionAssert.AreEqual(new[] { "/a.bas", "Sandbox/", "/b.cls" }, config.Sync.Ignore);
            CollectionAssert.IsEmpty(config.Errors);
            Assert.AreEqual(3, config.Indenter.IndentSpaces);
            StringAssert.Contains("// Files that are never synced", text);
            StringAssert.Contains("// How code is indented", text);
        }

        private static string AddTo(string json, params string[] patterns)
        {
            Assert.IsTrue(ProjectConfigWriter.TryAddToIgnoreList(json, patterns, out var updated), "the array should have been found");
            return updated;
        }

        private static string[] IgnoreOf(string json)
        {
            var dir = Path.Combine(Path.GetTempPath(), "rdign-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                File.WriteAllText(Path.Combine(dir, ProjectConfig.FileName), json);
                var config = new ProjectConfigReader(new FileSystem()).Read(dir, new IndenterSettings(true));
                CollectionAssert.IsEmpty(config.Errors, json);
                return config.Sync.Ignore.ToArray();
            }
            finally
            {
                Directory.Delete(dir, true);
            }
        }

        [TestCase("{ \"sync\": { \"ignore\": [] } }")]
        [TestCase("{ \"sync\": { \"ignore\": [\r\n  ] } }")]
        [TestCase("{\r\n  \"sync\": {\r\n    \"ignore\": []\r\n  }\r\n}")]
        [TestCase("{\n  \"sync\": {\n    \"ignore\": [\n      // nothing yet\n    ]\n  }\n}")]
        [Category("Settings")]
        public void AddsToAnEmptyArray(string json)
        {
            CollectionAssert.AreEqual(new[] { "/a", "/b" }, IgnoreOf(AddTo(json, "/a", "/b")));
        }

        [TestCase("{ \"sync\": { \"ignore\": [\"/x\"] } }")]
        [TestCase("{\r\n  \"sync\": {\r\n    \"ignore\": [\r\n      \"/x\"\r\n    ]\r\n  }\r\n}")]
        [TestCase("{\n  \"sync\": {\n    \"ignore\": [\n      \"/x\" // the first\n    ],\n    \"onViolation\": \"warn\"\n  }\n}")]
        [Category("Settings")]
        public void AddsAfterTheLastEntry(string json)
        {
            var updated = AddTo(json, "/a", "/b");

            CollectionAssert.AreEqual(new[] { "/x", "/a", "/b" }, IgnoreOf(updated));
            if (json.Contains("// the first"))
            {
                StringAssert.Contains("// the first", updated);
                StringAssert.Contains("\"warn\"", updated);
            }
        }

        [Test]
        [Category("Settings")]
        public void MultilineArrayGetsOneEntryPerLine()
        {
            var json = "{\r\n  \"sync\": {\r\n    \"ignore\": [\r\n      \"/x\"\r\n    ]\r\n  }\r\n}";

            var updated = AddTo(json, "/a");

            StringAssert.Contains("\"/x\",\r\n      \"/a\"\r\n    ]", updated);
        }

        [Test]
        [Category("Settings")]
        public void FileWithoutAnIgnoreArrayFallsBackToRewritingTheTree()
        {
            var path = Path.Combine(_root, ProjectConfig.FileName);
            File.WriteAllText(path, "{ \"indenter\": { \"indentSpaces\": 2 } }");

            new ProjectConfigWriter(new FileSystem()).AddIgnored(path, _root, new[] { "/a.bas" });
            var config = new ProjectConfigReader(new FileSystem()).Read(_root, new IndenterSettings(true));

            CollectionAssert.AreEqual(new[] { "/a.bas" }, config.Sync.Ignore);
            Assert.AreEqual(2, config.Indenter.IndentSpaces);
        }

        [Test]
        [Category("Settings")]
        public void SpecialCharactersInPatternsAreEscaped()
        {
            var updated = AddTo("{ \"sync\": { \"ignore\": [] } }", "/a \"b\"\\c");

            CollectionAssert.AreEqual(new[] { "/a \"b\"\\c" }, IgnoreOf(updated));
        }
    }
}
