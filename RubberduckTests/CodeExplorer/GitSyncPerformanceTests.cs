using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Abstractions;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.CodeExplorer.GitSync;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    [Category("GitSync")]
    public class SignatureCacheTests
    {
        private const string Text = "Attribute VB_Name = \"Foo\"\r\nPublic Sub A()\r\nDebug.Print 1\r\nEnd Sub\r\n";

        private static ISimpleIndenter IndenterFor(IIndenterSettings settings) => new Indenter(null, () => settings);

        private static string TempFile() => Path.Combine(Path.GetTempPath(), "rdcache-" + Guid.NewGuid().ToString("N") + ".json");

        [Test]
        public void SameTextIsComputedOnce()
        {
            var settings = new IndenterSettings(true);
            var cache = new SignatureCache(settings);

            var first = cache.Signature(Text, settings, IndenterFor(settings));
            var second = cache.Signature(Text, settings, IndenterFor(settings));

            Assert.AreEqual(first, second);
            Assert.AreEqual(1, cache.Misses);
            Assert.AreEqual(1, cache.Hits);
        }

        [Test]
        public void CachedSignatureEqualsTheComputedOne()
        {
            var settings = new IndenterSettings(true);

            Assert.AreEqual(
                SyncText.Signature(Text, settings, IndenterFor(settings)),
                new SignatureCache(settings).Signature(Text, settings, IndenterFor(settings)));
        }

        [Test]
        public void ChangedTextIsAMiss()
        {
            var settings = new IndenterSettings(true);
            var cache = new SignatureCache(settings);

            cache.Signature(Text, settings, IndenterFor(settings));
            cache.Signature(Text.Replace("Print 1", "Print 2"), settings, IndenterFor(settings));

            Assert.AreEqual(2, cache.Misses);
            Assert.AreEqual(0, cache.Hits);
        }

        [Test]
        public void ChangedSettingsDoNotReuseOldEntries()
        {
            var four = new IndenterSettings(true) { IndentSpaces = 4 };
            var two = new IndenterSettings(true) { IndentSpaces = 2 };
            var first = new SignatureCache(four);
            first.Signature(Text, four, IndenterFor(four));

            var path = TempFile();
            try
            {
                first.Save(new FileSystem(), path);
                var other = SignatureCache.Load(new FileSystem(), two, path);
                other.Signature(Text, two, IndenterFor(two));

                Assert.AreEqual(1, other.Misses, "an entry made under other settings must not be used");
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Test]
        public void SavedEntriesAreUsedAfterLoading()
        {
            var settings = new IndenterSettings(true);
            var path = TempFile();
            try
            {
                var cache = new SignatureCache(settings);
                var signature = cache.Signature(Text, settings, IndenterFor(settings));
                cache.Save(new FileSystem(), path);

                var loaded = SignatureCache.Load(new FileSystem(), settings, path);

                Assert.AreEqual(signature, loaded.Signature(Text, settings, IndenterFor(settings)));
                Assert.AreEqual(1, loaded.Hits);
                Assert.AreEqual(0, loaded.Misses);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Test]
        public void BrokenCacheFileStartsEmpty()
        {
            var settings = new IndenterSettings(true);
            var path = TempFile();
            try
            {
                File.WriteAllText(path, "{ not json");

                var cache = SignatureCache.Load(new FileSystem(), settings, path);
                cache.Signature(Text, settings, IndenterFor(settings));

                Assert.AreEqual(1, cache.Misses);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Test]
        public void SeveralThreadsGetTheSameAnswer()
        {
            var settings = new IndenterSettings(true);
            var cache = new SignatureCache(settings);
            var indenter = IndenterFor(settings);
            var expected = SyncText.Signature(Text, settings, indenter);

            var results = Enumerable.Range(0, 200).AsParallel().Select(_ => cache.Signature(Text, settings, indenter)).ToList();

            Assert.IsTrue(results.All(r => r == expected));
        }
    }

    [TestFixture]
    [Category("GitSync")]
    public class GitSyncPlannerParallelTests
    {
        private static ModuleInput Module(int i, string body) => new ModuleInput
        {
            Name = "Mod" + i,
            Extension = ".bas",
            ExportedText = $"Attribute VB_Name = \"Mod{i}\"\r\nPublic Sub A{i}()\r\n{body}\r\nEnd Sub\r\n"
        };

        //The same code with other line endings and indentation, which must not count as a difference.
        private static FileInput File(int i, string body) => new FileInput
        {
            Name = "Mod" + i,
            RelativePath = $"Mod{i}.bas",
            Text = $"Attribute VB_Name = \"Mod{i}\"\nPublic Sub A{i}()\n        {body}\nEnd Sub\n"
        };

        private static GitSyncPlanner PlannerFor(IIndenterSettings settings, SignatureCache cache = null) =>
            new GitSyncPlanner(settings, new Indenter(null, () => settings), new SyncConfig(), cache);

        [Test]
        public void ManyModulesAreComparedCorrectlyAndInOrder()
        {
            var settings = new IndenterSettings(true);
            var modules = Enumerable.Range(0, 300).Select(i => Module(i, i % 3 == 0 ? "Debug.Print 1" : "Debug.Print 2")).ToList();
            var files = Enumerable.Range(0, 300).Select(i => File(i, "Debug.Print 1")).ToList();

            var planner = PlannerFor(settings);
            var plan = planner.Plan(modules, files, new GitSyncState(), GitIgnoreSet.Empty);

            Assert.AreEqual(100, plan.InSyncCount, "every third module has the same code as its file");
            Assert.AreEqual(200, plan.Items.Count);
            CollectionAssert.AreEqual(
                plan.Items.Select(i => i.Name).OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList(),
                plan.Items.Select(i => i.Name).ToList());
            Assert.AreEqual(300, planner.Completed);
        }

        [Test]
        public void SecondRunWithTheSameCacheDoesNoIndenting()
        {
            var settings = new IndenterSettings(true);
            var cache = new SignatureCache(settings);
            var modules = Enumerable.Range(0, 50).Select(i => Module(i, "Debug.Print 1")).ToList();
            var files = Enumerable.Range(0, 50).Select(i => File(i, "Debug.Print 2")).ToList();

            PlannerFor(settings, cache).Plan(modules, files, new GitSyncState(), GitIgnoreSet.Empty);
            var missesAfterFirst = cache.Misses;
            PlannerFor(settings, cache).Plan(modules, files, new GitSyncState(), GitIgnoreSet.Empty);

            Assert.AreEqual(100, missesAfterFirst);
            Assert.AreEqual(100, cache.Misses, "the second run must come entirely from the cache");
        }

        [Test]
        public void CancellationStopsThePlan()
        {
            var settings = new IndenterSettings(true);
            var modules = Enumerable.Range(0, 100).Select(i => Module(i, "Debug.Print 1")).ToList();
            using (var cancelled = new CancellationTokenSource())
            {
                cancelled.Cancel();

                Assert.Throws<OperationCanceledException>(() =>
                    PlannerFor(settings).Plan(modules, new FileInput[0], new GitSyncState(), GitIgnoreSet.Empty, cancelled.Token));
            }
        }

        [Test]
        public void IgnoredItemsAreNotIndentedAtAll()
        {
            var settings = new IndenterSettings(true);
            var cache = new SignatureCache(settings);
            var ignored = new GitIgnoreSet(new[] { new KeyValuePair<string, GitIgnore>(string.Empty, new GitIgnore(new[] { "*.bas" })) });

            PlannerFor(settings, cache).Plan(new[] { Module(1, "Debug.Print 1") }, new[] { File(1, "Debug.Print 2") }, new GitSyncState(), ignored);

            Assert.AreEqual(0, cache.Misses);
        }
    }

    [TestFixture]
    [Category("GitSync")]
    public class GitSyncIndenterSettingsTests
    {
        [SetUp]
        public void SetUp() => GitSyncIndenterSettings.Forget();

        [TearDown]
        public void TearDown() => GitSyncIndenterSettings.Forget();

        [Test]
        public void ProjectWithoutAFolderGetsTheUsersSettings()
        {
            var user = new IndenterSettings(true);

            Assert.AreSame(user, GitSyncIndenterSettings.For(new FileSystem(), null, user));
            Assert.AreSame(user, GitSyncIndenterSettings.For(new FileSystem(), @"C:\no\such\book-" + Guid.NewGuid() + ".xlsm", user));
        }

        [Test]
        public void UnsavedWorkbooksNeverTouchTheFileSystem()
        {
            var user = new IndenterSettings(true);

            Assert.AreSame(user, GitSyncIndenterSettings.For(null, string.Empty, user));
        }
    }
}
