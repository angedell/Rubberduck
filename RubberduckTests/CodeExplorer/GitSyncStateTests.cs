using System;
using System.IO;
using System.IO.Abstractions;
using NUnit.Framework;
using Rubberduck.UI.CodeExplorer.GitSync;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    [Category("GitSync")]
    public class GitSyncStateTests
    {
        private string _root;

        [SetUp]
        public void SetUp()
        {
            _root = Path.Combine(Path.GetTempPath(), "RdGitSyncState", Path.GetRandomFileName());
            Directory.CreateDirectory(_root);
        }

        [TearDown]
        public void TearDown() => Directory.Delete(_root, true);

        [Test]
        public void StateLivesNextToTheConfig_OrInTheSyncFolder()
        {
            var store = new GitSyncStateStore(new FileSystem());

            Assert.AreEqual(Path.Combine(_root, GitSyncState.FileName), store.PathFor(Path.Combine(_root, "rubberduck.json"), @"C:\elsewhere"));
            Assert.AreEqual(Path.Combine(_root, GitSyncState.FileName), store.PathFor(null, _root));
        }

        [Test]
        public void SavedStateIsLoadedAgain()
        {
            var store = new GitSyncStateStore(new FileSystem());
            var path = Path.Combine(_root, GitSyncState.FileName);
            var state = new GitSyncState();
            state.Modules["Foo"] = new ModuleBaseline { Module = "a", File = "b", Path = "A/Foo.cls", Folder = "A" };

            store.Save(path, state);
            var loaded = store.Load(path).Modules["foo"];

            Assert.AreEqual("a", loaded.Module);
            Assert.AreEqual("b", loaded.File);
            Assert.AreEqual("A/Foo.cls", loaded.Path);
            Assert.AreEqual("A", loaded.Folder);
        }

        [Test]
        public void MissingOrBrokenFileGivesEmptyState()
        {
            var store = new GitSyncStateStore(new FileSystem());
            var path = Path.Combine(_root, GitSyncState.FileName);

            Assert.IsEmpty(store.Load(path).Modules);

            File.WriteAllText(path, "{ not json");
            Assert.IsEmpty(store.Load(path).Modules);
        }

        [Test]
        public void ModulesAreWrittenInAlphabeticalOrder_SoDiffsStayQuiet()
        {
            var store = new GitSyncStateStore(new FileSystem());
            var path = Path.Combine(_root, GitSyncState.FileName);
            var state = new GitSyncState();
            state.Modules["Zed"] = new ModuleBaseline();
            state.Modules["Alpha"] = new ModuleBaseline();

            store.Save(path, state);

            var text = File.ReadAllText(path);
            Assert.Less(text.IndexOf("Alpha", StringComparison.Ordinal), text.IndexOf("Zed", StringComparison.Ordinal));
        }
    }
}
