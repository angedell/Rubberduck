using System.Collections.Generic;
using System.IO;
using System.IO.Abstractions;
using System.Linq;
using NUnit.Framework;
using Rubberduck.UI.CodeExplorer.GitSync;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    [Category("GitSync")]
    public class GitIgnoreTests
    {
        private static GitIgnoreSet RootRules(params string[] lines)
        {
            return new GitIgnoreSet(new[] { new KeyValuePair<string, GitIgnore>(string.Empty, new GitIgnore(lines)) });
        }

        [Test]
        [TestCase("Legacy/", "Legacy", true, true)]
        [TestCase("Legacy/", "Legacy", false, false)]
        [TestCase("Legacy/", "Legacy/Foo.bas", false, true)]
        [TestCase("Legacy/", "Models/Legacy/Foo.bas", false, true)]
        [TestCase("/Legacy/", "Models/Legacy/Foo.bas", false, false)]
        [TestCase("/Legacy/", "Legacy/Foo.bas", false, true)]
        [TestCase("*.bas", "Models/Foo.bas", false, true)]
        [TestCase("*.bas", "Models/Foo.cls", false, false)]
        [TestCase("Models/Old", "Models/Old/Foo.bas", false, true)]
        [TestCase("Models/Old", "Other/Models/Old/Foo.bas", false, false)]
        [TestCase("**/Old/*.bas", "A/B/Old/Foo.bas", false, true)]
        [TestCase("**/Old/*.bas", "Old/Foo.bas", false, true)]
        [TestCase("Models/**", "Models/Sub/Foo.bas", false, true)]
        [TestCase("Foo?.bas", "Foo1.bas", false, true)]
        [TestCase("Foo[0-9].bas", "Foo7.bas", false, true)]
        [TestCase("Foo[!0-9].bas", "Foo7.bas", false, false)]
        [TestCase("# Legacy/", "Legacy/Foo.bas", false, false)]
        [TestCase("legacy/", "Legacy/Foo.bas", false, true)]
        public void IsIgnored(string rule, string path, bool isDirectory, bool expected)
        {
            Assert.AreEqual(expected, RootRules(rule).IsIgnored(path, isDirectory));
        }

        [Test]
        public void NegationReIncludesFile()
        {
            var rules = RootRules("*.bas", "!Keep.bas");
            Assert.IsTrue(rules.IsIgnored("Foo.bas", false));
            Assert.IsFalse(rules.IsIgnored("Keep.bas", false));
        }

        [Test]
        public void FileInsideIgnoredDirectoryCannotBeReIncluded()
        {
            var rules = RootRules("Legacy/", "!Legacy/Keep.bas");
            Assert.IsTrue(rules.IsIgnored("Legacy/Keep.bas", false));
        }

        [Test]
        public void NestedIgnoreFileAppliesOnlyBelowItsFolderAndWinsOverTheRoot()
        {
            var rules = new GitIgnoreSet(new[]
            {
                new KeyValuePair<string, GitIgnore>(string.Empty, new GitIgnore(new[] { "*.bas" })),
                new KeyValuePair<string, GitIgnore>("Models", new GitIgnore(new[] { "!Keep.bas", "Tmp/" }))
            });

            Assert.IsFalse(rules.IsIgnored("Models/Keep.bas", false));
            Assert.IsTrue(rules.IsIgnored("Models/Tmp/Foo.cls", false));
            Assert.IsFalse(rules.IsIgnored("Other/Tmp/Foo.cls", false));
            Assert.IsTrue(rules.IsIgnored("Other/Foo.bas", false));
        }

        [Test]
        public void BackslashSeparatorsAreAccepted()
        {
            Assert.IsTrue(RootRules("Legacy/").IsIgnored("Models\\Legacy\\Foo.bas", false));
        }
    }

    [TestFixture]
    [Category("GitSync")]
    public class GitSyncIgnoreIntegrationTests
    {
        [Test]
        public void SourceFiles_SkipsFoldersExcludedByIgnoreFile()
        {
            var root = Path.Combine(Path.GetTempPath(), "RdGitSyncTests", Path.GetRandomFileName());
            try
            {
                Directory.CreateDirectory(Path.Combine(root, "Keep"));
                Directory.CreateDirectory(Path.Combine(root, "Legacy", "Deep"));
                File.WriteAllText(Path.Combine(root, GitSyncFolderProcessor.IgnoreFileName), "Legacy/\r\n");
                File.WriteAllText(Path.Combine(root, "Keep", "A.bas"), "x");
                File.WriteAllText(Path.Combine(root, "Legacy", "B.bas"), "x");
                File.WriteAllText(Path.Combine(root, "Legacy", "Deep", "C.bas"), "x");

                var processor = new GitSyncFolderProcessor(new FileSystem());
                var files = processor.SourceFiles(root).Select(Path.GetFileName).ToList();

                CollectionAssert.AreEqual(new[] { "A.bas" }, files);
                Assert.IsTrue(processor.IsIgnored(root, Path.Combine(root, "Legacy", "B.bas")));
                Assert.IsFalse(processor.IsIgnored(root, Path.Combine(root, "Keep", "A.bas")));
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }
    }
}
