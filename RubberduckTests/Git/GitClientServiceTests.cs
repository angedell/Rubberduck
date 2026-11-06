using System;
using System.IO;
using System.IO.Abstractions;
using System.Linq;
using System.Text;
using Moq;
using NUnit.Framework;
using Rubberduck.Git;
using Rubberduck.Settings;
using Rubberduck.UI.CodeExplorer.GitSync;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Git
{
    [TestFixture]
    public class GitClientServiceTests
    {
        private string _folder;
        private ProjectRepositories _repositories;

        [SetUp]
        public void SetUp()
        {
            _folder = Path.Combine(Path.GetTempPath(), "rd-client-" + Guid.NewGuid());
            Directory.CreateDirectory(_folder);
            _repositories = new ProjectRepositories(new FileSystem(), Path.Combine(_folder, "map.json"), Path.Combine(_folder, "repos"));
        }

        [TearDown]
        public void TearDown()
        {
            foreach (var file in Directory.GetFiles(_folder, "*", SearchOption.AllDirectories))
            {
                File.SetAttributes(file, FileAttributes.Normal);
            }

            Directory.Delete(_folder, true);
        }

        private string Workbook(string name)
        {
            var path = Path.Combine(_folder, name);
            File.WriteAllText(path, "workbook");
            return path;
        }

        // ---- ProjectRepositories ----

        [Test]
        public void Resolve_GivesANewWorkbookARepositoryAndRemembersIt()
        {
            var first = _repositories.Resolve(Workbook("a.xlsm"), "Alpha");
            var again = _repositories.Resolve(Workbook("a.xlsm"), "Alpha");

            Assert.AreEqual(ProjectRepositoryOutcome.New, first.Outcome);
            Assert.AreEqual(ProjectRepositoryOutcome.Known, again.Outcome);
            Assert.AreEqual(first.Id, again.Id);
        }

        [Test]
        public void Resolve_WhenTheOldFileIsGone_FollowsTheWorkbookToItsNewPath()
        {
            var old = Workbook("old.xlsm");
            var id = _repositories.Resolve(old, "Alpha").Id;
            File.Move(old, Path.Combine(_folder, "renamed.xlsm"));

            var moved = _repositories.Resolve(Path.Combine(_folder, "renamed.xlsm"), "Alpha");

            Assert.AreEqual(ProjectRepositoryOutcome.Moved, moved.Outcome);
            Assert.AreEqual(id, moved.Id);
            Assert.AreEqual(ProjectRepositoryOutcome.Known, _repositories.Resolve(Path.Combine(_folder, "renamed.xlsm"), "Alpha").Outcome);
        }

        [Test]
        public void Resolve_WhenTheOldFileIsStillThere_ReportsACopyAndRegistersNothing()
        {
            var original = Workbook("original.xlsm");
            var id = _repositories.Resolve(original, "Alpha").Id;

            var copy = _repositories.Resolve(Workbook("copy.xlsm"), "Alpha");

            Assert.AreEqual(ProjectRepositoryOutcome.CopyOf, copy.Outcome);
            Assert.AreEqual(id, copy.CopyOfId);
            Assert.AreEqual(original, copy.CopyOfPath);
            Assert.IsNull(copy.Id);
        }

        [Test]
        public void Adopt_CanContinueTheOriginalsHistoryOrStartEmpty()
        {
            var id = _repositories.Resolve(Workbook("original.xlsm"), "Alpha").Id;
            Directory.CreateDirectory(_repositories.RootOf(id));
            File.WriteAllText(Path.Combine(_repositories.RootOf(id), "marker.txt"), "history");

            var forked = _repositories.Adopt(Workbook("fork.xlsm"), "Alpha", id, true);
            var fresh = _repositories.Adopt(Workbook("fresh.xlsm"), "Alpha", id, false);

            Assert.IsTrue(File.Exists(Path.Combine(_repositories.RootOf(forked), "marker.txt")));
            Assert.IsFalse(Directory.Exists(_repositories.RootOf(fresh)));
            Assert.AreEqual(3, new[] { id, forked, fresh }.Distinct().Count());
        }

        [Test]
        public void Resolve_KeepsAMapItCannotReadInsteadOfOverwritingIt()
        {
            var map = Path.Combine(_folder, "map.json");
            File.WriteAllText(map, "{ not json");

            var resolution = _repositories.Resolve(Workbook("a.xlsm"), "Alpha");

            Assert.AreEqual(ProjectRepositoryOutcome.New, resolution.Outcome);
            Assert.AreEqual(1, Directory.GetFiles(_folder, "map.json.damaged-*").Length);
        }

        // ---- GitClientService, with a real git and a snapshot that writes made-up modules ----

        private class FakeSnapshot : IProjectSnapshot
        {
            public string Module = "Option Explicit\r\n";
            public string Error;
            public int Taken;

            public string Imported;

            public GitSyncResult Import(IVBProject project, string tree, string moduleName, ISyncProgress progress = null)
            {
                Imported = moduleName + "=" + File.ReadAllText(Path.Combine(tree, "Module1.bas"));
                return new GitSyncResult();
            }

            public GitSyncResult ImportFiles(IVBProject project, string tree, System.Collections.Generic.IReadOnlyCollection<string> relativePaths, ISyncProgress progress = null)
            {
                ImportedFiles = relativePaths.ToArray();
                return new GitSyncResult();
            }

            public string[] ImportedFiles;

            public GitSyncResult Take(IVBProject project, string tree, ISyncProgress progress = null)
            {
                Taken++;
                File.WriteAllText(Path.Combine(tree, "Module1.bas"), Module, Encoding.ASCII);
                //What Git sync keeps in the tree, and the client reads to find a module's file.
                File.WriteAllText(Path.Combine(tree, "rubberduck.sync.json"), "{\"Version\":1,\"Modules\":{\"Module1\":{\"Path\":\"Module1.bas\"}}}");
                var result = new GitSyncResult();
                if (Error != null)
                {
                    result.Errors.Add(Error);
                }

                return result;
            }
        }

        private static IVBProject Project(string path, string name)
        {
            var project = new Mock<IVBProject>();
            project.SetupGet(p => p.FileName).Returns(path);
            project.SetupGet(p => p.Name).Returns(name);
            return project.Object;
        }

        private (GitClientService service, FakeSnapshot snapshot) Service(bool enabled = true, bool gitInstalled = true)
        {
            var installation = new GitLocator(new FileSystem(), GitLocator.RunVersion, null, Environment.GetEnvironmentVariable("PATH"), GitLocator.DefaultOwnFolder,
                Environment.GetEnvironmentVariable("ProgramFiles")).Find();
            if (installation == null)
            {
                Assert.Ignore("No git on this machine.");
            }

            var provisioner = new Mock<IGitProvisioner>();
            provisioner.Setup(p => p.Find(It.IsAny<string>())).Returns(gitInstalled ? installation : null);
            var snapshot = new FakeSnapshot();
            var service = new GitClientService(provisioner.Object, new FileSystem(), snapshot, _repositories, () => new GeneralSettings { GitClientEnabled = enabled }, new MemoryCredentials());
            return (service, snapshot);
        }

        private class MemoryCredentials : ICredentialStore
        {
            private readonly System.Collections.Generic.Dictionary<string, StoredCredential> _saved = new System.Collections.Generic.Dictionary<string, StoredCredential>();

            public StoredCredential Read(string target) => _saved.TryGetValue(target, out var credential) ? credential : null;
            public void Write(string target, string user, string secret) => _saved[target] = new StoredCredential(user, secret);
            public bool Delete(string target) => _saved.Remove(target);
            public System.Collections.Generic.IReadOnlyList<string> List(string prefix) => _saved.Keys.Where(key => key.StartsWith(prefix)).ToList();
        }

        [TestCase("git@github.com:o/r.git")]
        [TestCase("http://github.com/o/r.git")]
        [TestCase(@"C:\repos\r.git")]
        [TestCase("https://user:pw@github.com/o/r.git")]
        public void SetRemote_RefusesAnythingButAPlainHttpsAddress(string url)
        {
            var (service, _) = Service();

            Assert.Throws<GitClientException>(() => service.SetRemote(Project(Workbook("a.xlsm"), "Alpha"), url));
        }

        [Test]
        public void SetRemote_KeepsTheAddressWithTheProject()
        {
            var (service, _) = Service();
            var project = Project(Workbook("a.xlsm"), "Alpha");

            service.SetRemote(project, " https://github.com/o/r.git ");

            Assert.AreEqual("https://github.com/o/r.git", service.RemoteUrl(project));
            Assert.AreEqual("github.com", service.RemoteHost(project));
        }

        [Test]
        public void Push_WithoutARemote_SaysSo()
        {
            var (service, _) = Service();
            var project = Project(Workbook("a.xlsm"), "Alpha");
            service.Commit(project, "first");

            StringAssert.Contains("no remote", Assert.Throws<GitClientException>(() => service.Push(project)).Message);
        }

        [Test]
        public void Pull_RefusesWhileTheProjectHasChangesThatAreNotCommitted()
        {
            var (service, snapshot) = Service();
            var project = Project(Workbook("a.xlsm"), "Alpha");
            service.Commit(project, "first");
            service.SetRemote(project, "https://github.com/o/r.git");
            snapshot.Module = "Option Explicit\r\n'edited\r\n";

            StringAssert.Contains("not committed", Assert.Throws<GitClientException>(() => service.Pull(project)).Message);
        }

        [Test]
        public void Tokens_AreSavedPerHostAndCanBeForgotten()
        {
            var (service, _) = Service();

            Assert.IsFalse(service.HasToken("github.com"));
            service.SaveToken("github.com", "angelo", "ghp_x");
            Assert.IsTrue(service.HasToken("github.com"));
            Assert.IsFalse(service.HasToken("gitlab.com"));
            Assert.IsTrue(service.ForgetToken("github.com"));
            Assert.IsFalse(service.HasToken("github.com"));
            Assert.Throws<GitClientException>(() => service.SaveToken("github.com", "angelo", "  "));
        }

        [Test]
        public void Commit_SavesTheProjectAndReportsNothingToCommitTheSecondTime()
        {
            var (service, _) = Service();
            var project = Project(Workbook("a.xlsm"), "Alpha");

            var first = service.Commit(project, "first");
            var second = service.Commit(project, "again");

            Assert.IsNotNull(first.CommitId);
            Assert.IsNull(second.CommitId);
            Assert.AreEqual(new[] { "first" }, service.History(project).Select(commit => commit.Subject).ToArray());
        }

        [Test]
        public void Changes_ShowWhatTheProjectChangedSinceTheLastCommit()
        {
            var (service, snapshot) = Service();
            var project = Project(Workbook("a.xlsm"), "Alpha");
            service.Commit(project, "first");
            snapshot.Module = "Option Explicit\r\nPublic Const X = 1\r\n";

            var changes = service.Changes(project);

            Assert.AreEqual(1, changes.Count);
            Assert.AreEqual("Module1.bas", changes[0].Path);
            Assert.AreEqual(GitChangeKind.Modified, changes[0].Kind);
        }

        [Test]
        public void Commit_WhenAModuleCouldNotBeWritten_CommitsNothing()
        {
            var (service, snapshot) = Service();
            snapshot.Error = "Could not sync 'Module1': locked";
            var project = Project(Workbook("a.xlsm"), "Alpha");

            var exception = Assert.Throws<GitClientException>(() => service.Commit(project, "x"));

            StringAssert.Contains("locked", exception.Message);
            Assert.AreEqual(0, service.History(project).Count);
        }

        [Test]
        public void Commit_OfACopiedWorkbook_AsksUnlessToldWhatToDo()
        {
            var (service, _) = Service();
            service.Commit(Project(Workbook("original.xlsm"), "Alpha"), "first");
            var copy = Project(Workbook("copy.xlsm"), "Alpha");

            var asked = Assert.Throws<ProjectCopiedException>(() => service.Commit(copy, "x"));
            service.Commit(copy, "continued", continueCopiedHistory: true);

            Assert.AreEqual(Path.Combine(_folder, "original.xlsm"), asked.CopyOfPath);
            Assert.AreEqual(new[] { "first" }, service.History(copy).Select(commit => commit.Subject).Reverse().ToArray());
        }

        [Test]
        public void ModuleHistory_ListsTheCommitsThatChangedTheModule()
        {
            var (service, snapshot) = Service();
            var project = Project(Workbook("a.xlsm"), "Alpha");
            service.Commit(project, "first");
            snapshot.Module = "Option Explicit\r\n' v2\r\n";
            service.Commit(project, "second");

            var history = service.ModuleHistory(project, "Module1");

            CollectionAssert.AreEqual(new[] { "second", "first" }, history.Select(commit => commit.Subject).ToArray());
            Assert.IsTrue(history.All(commit => commit.Path == "Module1.bas"));
        }

        [Test]
        public void ModuleHistory_OfAModuleThatWasNeverCommitted_SaysSo()
        {
            var (service, _) = Service();
            var project = Project(Workbook("a.xlsm"), "Alpha");
            service.Commit(project, "first");

            var exception = Assert.Throws<GitClientException>(() => service.ModuleHistory(project, "Unknown"));

            StringAssert.Contains("not been committed", exception.Message);
        }

        [Test]
        public void RestoreModule_WritesTheOldVersionIntoTheTreeAndImportsIt_WithoutCommitting()
        {
            var (service, snapshot) = Service();
            var project = Project(Workbook("a.xlsm"), "Alpha");
            service.Commit(project, "first");
            snapshot.Module = "Option Explicit\r\n' v2\r\n";
            service.Commit(project, "second");
            var first = service.ModuleHistory(project, "Module1").Last();

            service.RestoreModule(project, "Module1", first);

            Assert.AreEqual("Module1=Option Explicit\r\n", snapshot.Imported);
            Assert.AreEqual(2, service.History(project).Count);
        }

        [Test]
        public void Commit_RefusesWhenTheClientIsOffGitIsMissingOrTheWorkbookIsUnsaved()
        {
            Assert.Throws<GitClientException>(() => Service(enabled: false).service.Commit(Project(Workbook("a.xlsm"), "A"), "x"));
            Assert.Throws<GitClientException>(() => Service(gitInstalled: false).service.Commit(Project(Workbook("b.xlsm"), "B"), "x"));

            var (service, snapshot) = Service();
            Assert.Throws<GitClientException>(() => service.Commit(Project(null, "Unsaved"), "x"));
            Assert.AreEqual(0, snapshot.Taken);
        }
    }
}
