using System;
using System.IO;
using System.IO.Abstractions;
using NUnit.Framework;
using Rubberduck.UI.CodeExplorer.GitSync;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    [Category("GitSync")]
    public class SyncDirectionResolverTests
    {
        private static readonly DateTime Now = new DateTime(2025, 1, 1, 12, 0, 0, DateTimeKind.Utc);

        private static ModuleSnapshot Snapshot(DateTime? moduleChanged = null)
        {
            return new ModuleSnapshot { ModuleHash = "m1", FileHash = "f1", ModuleChangedUtc = moduleChanged };
        }

        private static SyncDecision Decide(ModuleSnapshot snapshot, string module, string file, bool equal = false, DateTime? fileWritten = null)
        {
            return SyncDirectionResolver.Decide(snapshot, module, file, () => equal, fileWritten ?? Now, Now);
        }

        [Test]
        public void NothingChanged_DoesNothing()
        {
            Assert.AreEqual(SyncDirection.None, Decide(Snapshot(), "m1", "f1").Direction);
        }

        [Test]
        public void OnlyModuleChanged_Exports()
        {
            var decision = Decide(Snapshot(), "m2", "f1");
            Assert.AreEqual(SyncDirection.Export, decision.Direction);
            Assert.IsFalse(decision.IsConflict);
        }

        [Test]
        public void OnlyFileChanged_Imports()
        {
            var decision = Decide(Snapshot(), "m1", "f2");
            Assert.AreEqual(SyncDirection.Import, decision.Direction);
            Assert.IsFalse(decision.IsConflict);
        }

        [Test]
        public void BothChangedToTheSameContent_OnlyRecords()
        {
            Assert.AreEqual(SyncDirection.Record, Decide(Snapshot(), "m2", "f2", equal: true).Direction);
        }

        [Test]
        public void BothChanged_NewerFileWins()
        {
            var decision = Decide(Snapshot(Now.AddMinutes(-10)), "m2", "f2", fileWritten: Now);
            Assert.AreEqual(SyncDirection.Import, decision.Direction);
            Assert.IsTrue(decision.IsConflict);
        }

        [Test]
        public void BothChanged_NewerModuleWins()
        {
            var decision = Decide(Snapshot(Now), "m2", "f2", fileWritten: Now.AddMinutes(-10));
            Assert.AreEqual(SyncDirection.Export, decision.Direction);
            Assert.IsTrue(decision.IsConflict);
        }

        [Test]
        public void BothChanged_ModuleChangeNotRecordedYet_CountsAsJustNow()
        {
            var decision = Decide(Snapshot(), "m2", "f2", fileWritten: Now.AddMinutes(-10));
            Assert.AreEqual(SyncDirection.Export, decision.Direction);
        }

        [Test]
        public void NeverSynced_ModuleOnly_Exports()
        {
            Assert.AreEqual(SyncDirection.Export, Decide(null, "m1", null).Direction);
        }

        [Test]
        public void NeverSynced_FileOnly_Imports()
        {
            Assert.AreEqual(SyncDirection.Import, Decide(null, null, "f1").Direction);
        }

        [Test]
        public void NeverSynced_BothDiffer_AsksWhoWins()
        {
            var decision = Decide(null, "m1", "f1");
            Assert.AreEqual(SyncDirection.Ask, decision.Direction);
            Assert.IsFalse(decision.IsConflict);
        }

        [Test]
        public void NeverSynced_BothEqual_OnlyRecords()
        {
            Assert.AreEqual(SyncDirection.Record, Decide(null, "m1", "f1", equal: true).Direction);
        }

        [Test]
        public void FileDeletedAfterSync_ModuleIsExportedAgain()
        {
            Assert.AreEqual(SyncDirection.Export, Decide(Snapshot(), "m1", null).Direction);
        }
    }

    [TestFixture]
    [Category("GitSync")]
    public class GitSyncManifestStoreTests
    {
        private string _file;

        [SetUp]
        public void SetUp()
        {
            _file = Path.Combine(Path.GetTempPath(), "RdGitSyncTests", Path.GetRandomFileName() + ".json");
        }

        [TearDown]
        public void TearDown()
        {
            if (File.Exists(_file))
            {
                File.Delete(_file);
            }
        }

        [Test]
        public void SavedSnapshotsAreLoadedAgain()
        {
            var store = new GitSyncManifestStore(new FileSystem(), _file);
            var manifest = new ProjectManifest { Root = @"C:\Sync" };
            manifest.Modules["Foo"] = new ModuleSnapshot { ModuleHash = "m", FileHash = "f", RelativePath = @"A\Foo.bas", Folder = "A" };
            store.Save("Book.xlsm", manifest);

            var loaded = new GitSyncManifestStore(new FileSystem(), _file).Load("Book.xlsm", @"C:\Sync");

            Assert.AreEqual("m", loaded.Modules["foo"].ModuleHash);
            Assert.AreEqual(@"A\Foo.bas", loaded.Modules["Foo"].RelativePath);
        }

        [Test]
        public void SnapshotsAreIgnoredWhenTheFolderChanged()
        {
            var store = new GitSyncManifestStore(new FileSystem(), _file);
            var manifest = new ProjectManifest { Root = @"C:\Sync" };
            manifest.Modules["Foo"] = new ModuleSnapshot { ModuleHash = "m" };
            store.Save("Book.xlsm", manifest);

            Assert.IsEmpty(store.Load("Book.xlsm", @"C:\Elsewhere").Modules);
        }

        [Test]
        public void UnsavedProjectsAreNotPersisted()
        {
            var store = new GitSyncManifestStore(new FileSystem(), _file);
            store.Save(null, new ProjectManifest { Root = @"C:\Sync" });

            Assert.IsFalse(File.Exists(_file));
        }
    }
}
