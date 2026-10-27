using System;
using System.IO;
using System.IO.Abstractions;
using NUnit.Framework;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;

namespace RubberduckTests.Settings
{
    [TestFixture]
    public class ProjectConfigTests
    {
        private string _root;

        [SetUp]
        public void SetUp()
        {
            _root = Path.Combine(Path.GetTempPath(), "rdconfig-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(_root, "src"));
        }

        [TearDown]
        public void TearDown()
        {
            Directory.Delete(_root, true);
        }

        private ProjectConfig Read(string json, IIndenterSettings baseSettings = null)
        {
            if (json != null)
            {
                File.WriteAllText(Path.Combine(_root, ProjectConfig.FileName), json);
            }

            return new ProjectConfigReader(new FileSystem()).Read(Path.Combine(_root, "src"), baseSettings ?? new IndenterSettings(true));
        }

        [Test]
        [Category("Settings")]
        public void NoFileMeansNotFound()
        {
            var config = Read(null);

            Assert.IsFalse(config.Found);
            Assert.IsNull(config.Indenter);
        }

        [Test]
        [Category("Settings")]
        public void FileInParentFolderIsFound()
        {
            var config = Read("{ }");

            Assert.IsTrue(config.Found);
            Assert.AreEqual(Path.Combine(_root, ProjectConfig.FileName), config.Path);
        }

        [Test]
        [Category("Settings")]
        public void FileOverridesOnlyWhatItMentions()
        {
            var user = new IndenterSettings(true) { IndentSpaces = 8, IndentCase = true };

            var config = Read(@"{ ""indenter"": { ""indentSpaces"": 2 } }", baseSettings: user);

            Assert.AreEqual(2, config.Indenter.IndentSpaces);
            Assert.IsTrue(config.Indenter.IndentCase);
            Assert.AreEqual(8, user.IndentSpaces, "the user's own settings must not change");
            Assert.IsEmpty(config.Errors);
        }

        [Test]
        [Category("Settings")]
        public void KeysAreCaseInsensitiveAndEnumsAreStrings()
        {
            var config = Read(@"{ ""indenter"": { ""IndentCase"": true, ""endOfLineCommentStyle"": ""samegap"" } }");

            Assert.IsTrue(config.Indenter.IndentCase);
            Assert.AreEqual(EndOfLineCommentStyle.SameGap, config.Indenter.EndOfLineCommentStyle);
            Assert.IsEmpty(config.Errors);
        }

        [Test]
        [Category("Settings")]
        public void SpecificForceSettingSurvivesTheGeneralOne()
        {
            var config = Read(@"{ ""indenter"": { ""forceStopInColumn1"": true, ""forceDebugStatementsInColumn1"": false } }");

            Assert.IsTrue(config.Indenter.ForceStopInColumn1);
        }

        [Test]
        [Category("Settings")]
        public void UnknownKeysAndBadValuesAreReported()
        {
            var config = Read(@"{ ""indenter"": { ""indentSpaces"": ""many"", ""nope"": 1 }, ""formatter"": {} }", baseSettings: new IndenterSettings(true) { IndentSpaces = 3 });

            Assert.AreEqual(3, config.Errors.Count);
            Assert.AreEqual(3, config.Indenter.IndentSpaces);
        }

        [Test]
        [Category("Settings")]
        public void MalformedJsonIsReportedAndKeepsBaseSettings()
        {
            var config = Read("{ not json", baseSettings: new IndenterSettings(true) { IndentSpaces = 3 });

            Assert.IsTrue(config.Found);
            Assert.AreEqual(1, config.Errors.Count);
            Assert.AreEqual(3, config.Indenter.IndentSpaces);
        }

        [Test]
        [Category("Settings")]
        public void SyncSectionIsRead()
        {
            var config = Read(@"{ ""sync"": { ""ignore"": [""/modScratch.bas"", ""Sandbox/""], ""enforce"": [""moduleWithoutFolder"", ""MissingAttribute""], ""onViolation"": ""warn"" } }");

            CollectionAssert.AreEqual(new[] { "/modScratch.bas", "Sandbox/" }, config.Sync.Ignore);
            CollectionAssert.AreEqual(new[] { "ModuleWithoutFolder", "MissingAttribute" }, config.Sync.Enforce);
            Assert.AreEqual(ViolationMode.Warn, config.Sync.OnViolation);
            Assert.IsEmpty(config.Errors);
        }

        [Test]
        [Category("Settings")]
        public void SyncDefaultsToFixingAndNothingEnforced()
        {
            var config = Read("{ }");

            Assert.IsEmpty(config.Sync.Enforce);
            Assert.IsEmpty(config.Sync.Ignore);
            Assert.AreEqual(ViolationMode.Fix, config.Sync.OnViolation);
        }

        [Test]
        [Category("Settings")]
        public void BadSyncSettingsAreReported()
        {
            var config = Read(@"{ ""sync"": { ""enforce"": [""NoSuchRule""], ""onViolation"": ""repair"", ""ignore"": ""x"", ""nope"": 1 } }");

            Assert.AreEqual(4, config.Errors.Count);
            Assert.IsEmpty(config.Sync.Enforce);
        }

        [Test]
        [Category("Settings")]
        public void WriterAddsIgnoredPatternsAndKeepsTheRest()
        {
            var path = Path.Combine(_root, ProjectConfig.FileName);
            File.WriteAllText(path, @"{ ""indenter"": { ""indentSpaces"": 2 }, ""sync"": { ""ignore"": [""/a.bas""] } }");

            new ProjectConfigWriter(new FileSystem()).AddIgnored(path, _root, new[] { "/a.bas", "/b.cls" });
            var config = new ProjectConfigReader(new FileSystem()).Read(Path.Combine(_root, "src"), new IndenterSettings(true));

            CollectionAssert.AreEqual(new[] { "/a.bas", "/b.cls" }, config.Sync.Ignore);
            Assert.AreEqual(2, config.Indenter.IndentSpaces);
            Assert.IsEmpty(config.Errors);
        }

        [Test]
        [Category("Settings")]
        public void WriterCreatesTheFileIfThereIsNone()
        {
            var created = new ProjectConfigWriter(new FileSystem()).AddIgnored(null, _root, new[] { "/a.bas" });

            Assert.AreEqual(Path.Combine(_root, ProjectConfig.FileName), created);
            CollectionAssert.AreEqual(new[] { "/a.bas" }, new ProjectConfigReader(new FileSystem()).Read(_root, new IndenterSettings(true)).Sync.Ignore);
        }
    }
}
