using System;
using System.IO;
using System.IO.Abstractions;
using System.Linq;
using System.Text;
using NUnit.Framework;
using Rubberduck.Git;

namespace RubberduckTests.Git
{
    /// <summary>Runs a real git, because what these protect against (a .git in the work tree, hooks, line endings) is git's behavior, not ours.</summary>
    [TestFixture]
    public class HiddenRepositoryTests
    {
        private string _folder;
        private string _gitPath;
        private GitRunner _git;
        private HiddenRepository _repository;

        [SetUp]
        public void SetUp()
        {
            var installation = new GitLocator(new FileSystem(), GitLocator.RunVersion, null, Environment.GetEnvironmentVariable("PATH"), GitLocator.DefaultOwnFolder,
                Environment.GetEnvironmentVariable("ProgramFiles")).Find();
            if (installation == null)
            {
                Assert.Ignore("No git on this machine.");
            }

            _folder = Path.Combine(Path.GetTempPath(), "rd-repo-" + Guid.NewGuid());
            _gitPath = installation.Path;
            _git = new GitRunner(_gitPath);
            _repository = new HiddenRepository(_git, new FileSystem(), _folder);
        }

        [TearDown]
        public void TearDown()
        {
            if (_folder != null && Directory.Exists(_folder))
            {
                foreach (var file in Directory.GetFiles(_folder, "*", SearchOption.AllDirectories))
                {
                    File.SetAttributes(file, FileAttributes.Normal);
                }

                Directory.Delete(_folder, true);
            }
        }

        private void Write(string relative, byte[] bytes)
        {
            var path = Path.Combine(_repository.WorkTree, relative);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllBytes(path, bytes);
        }

        private string Commit(string message = "test") => _repository.Commit(message, "Tester", "tester@example.test");

        [Test]
        public void EnsureCreated_PutsNothingGitRelatedInTheWorkTree()
        {
            _repository.EnsureCreated();
            Write("Module1.bas", Encoding.ASCII.GetBytes("Attribute VB_Name = \"Module1\"\r\n"));
            Commit();

            Assert.IsTrue(_repository.Exists);
            Assert.IsFalse(Directory.Exists(Path.Combine(_repository.WorkTree, ".git")));
            Assert.IsFalse(File.Exists(Path.Combine(_repository.WorkTree, ".git")));
            Assert.AreEqual(new[] { "Module1.bas" }, Directory.GetFileSystemEntries(_repository.WorkTree).Select(Path.GetFileName).ToArray());
        }

        private string[] Tracked() => _git.RunOrThrow(new[] { "--git-dir=" + _repository.GitDirectory, "ls-tree", "-r", "--name-only", "HEAD" }).Output
            .Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(line => line.Trim()).ToArray();

        [Test]
        public void GitSyncsBookkeeping_IsNeverCommittedNorListedAsAChange()
        {
            _repository.EnsureCreated();
            Write("M.bas", new byte[] { 65 });
            Write("rubberduck.json", new byte[] { 123, 125 });
            Write("rubberduck.sync.json", new byte[] { 1 });
            Commit("first");

            Write("rubberduck.sync.json", new byte[] { 2 });
            Write("M.bas", new byte[] { 66 });

            Assert.AreEqual(new[] { "M.bas" }, _repository.Status().Select(change => change.Path).ToArray());
            Commit("second");
            CollectionAssert.AreEquivalent(new[] { "M.bas", "rubberduck.json" }, Tracked());
            Assert.IsTrue(File.Exists(Path.Combine(_repository.WorkTree, "rubberduck.sync.json")));
        }

        [Test]
        public void EnsureCreated_TakesBookkeepingOutOfARepositoryThatAlreadyHoldsIt()
        {
            _repository.EnsureCreated();
            Write("M.bas", new byte[] { 65 });
            Commit("first");
            Write("rubberduck.sync.json", new byte[] { 1 });
            _git.RunOrThrow(new[] { "--git-dir=" + _repository.GitDirectory, "--work-tree=" + _repository.WorkTree, "add", "--force", "rubberduck.sync.json" });
            _git.RunOrThrow(new[] { "--git-dir=" + _repository.GitDirectory, "-c", "user.name=t", "-c", "user.email=t@t", "commit", "--quiet", "--message", "old" });
            CollectionAssert.Contains(Tracked(), "rubberduck.sync.json");

            _repository.EnsureCreated();
            Write("rubberduck.sync.json", new byte[] { 2 });
            Write("M.bas", new byte[] { 66 });

            Assert.AreEqual(new[] { "M.bas" }, _repository.Status().Select(change => change.Path).ToArray());
            Commit("after");
            CollectionAssert.DoesNotContain(Tracked(), "rubberduck.sync.json");
            Assert.IsTrue(File.Exists(Path.Combine(_repository.WorkTree, "rubberduck.sync.json")));
        }

        [Test]
        public void Status_ListsNewModifiedAndDeletedFiles()
        {
            _repository.EnsureCreated();
            Write("a.bas", new byte[] { 65 });
            Write("b.bas", new byte[] { 66 });
            Commit();

            Write("a.bas", new byte[] { 65, 65 });
            File.Delete(Path.Combine(_repository.WorkTree, "b.bas"));
            Write("sub dir/c.cls", new byte[] { 67 });

            var changes = _repository.Status().ToDictionary(change => change.Path, change => change.Kind);

            Assert.AreEqual(GitChangeKind.Modified, changes["a.bas"]);
            Assert.AreEqual(GitChangeKind.Deleted, changes["b.bas"]);
            Assert.AreEqual(GitChangeKind.Added, changes["sub dir/c.cls"]);
        }

        [Test]
        public void Commit_WithNothingChanged_ReturnsNull()
        {
            _repository.EnsureCreated();
            Write("a.bas", new byte[] { 65 });
            Assert.IsNotNull(Commit());

            Assert.IsNull(Commit());
            Assert.AreEqual(1, _repository.Log().Count);
        }

        [Test]
        public void Log_ListsTheNewestFirstAndIsEmptyBeforeTheFirstCommit()
        {
            _repository.EnsureCreated();
            Assert.AreEqual(0, _repository.Log().Count);

            Write("a.bas", new byte[] { 65 });
            Commit("first");
            Write("a.bas", new byte[] { 66 });
            var second = Commit("second: with \"quotes\"");

            var log = _repository.Log();

            Assert.AreEqual(new[] { "second: with \"quotes\"", "first" }, log.Select(commit => commit.Subject).ToArray());
            Assert.AreEqual(second, log[0].Id);
            Assert.AreEqual("Tester", log[0].Author);
        }

        [Test]
        public void ModulesComeBackByteForByte_CrLfAndAnsiIncluded()
        {
            //A CRLF module with a character from the ANSI code page: neither the line ending nor the byte may change.
            var bytes = Encoding.GetEncoding(1252).GetBytes("Attribute VB_Name = \"M\"\r\n' caf\u00e9\r\nSub A()\r\nEnd Sub\r\n");
            _repository.EnsureCreated();
            Write("M.bas", bytes);
            var id = Commit();

            Assert.IsEmpty(_repository.Status());
            Assert.AreEqual(bytes, _repository.ReadFileAt(id, "M.bas"));
        }

        [Test]
        public void RestoreFile_PutsTheOldVersionBack()
        {
            _repository.EnsureCreated();
            Write("a.bas", new byte[] { 1 });
            var first = Commit("first");
            Write("a.bas", new byte[] { 2 });
            Commit("second");

            _repository.RestoreFile(first, "a.bas");

            Assert.AreEqual(new byte[] { 1 }, File.ReadAllBytes(Path.Combine(_repository.WorkTree, "a.bas")));
        }

        [Test]
        public void Hooks_NeverRun()
        {
            _repository.EnsureCreated();
            var marker = Path.Combine(_folder, "hook-ran.txt");
            var hooks = Path.Combine(_repository.GitDirectory, "hooks");
            Directory.CreateDirectory(hooks);
            File.WriteAllText(Path.Combine(hooks, "pre-commit"), "#!/bin/sh\necho ran > \"" + marker.Replace('\\', '/') + "\"\n");
            Write("a.bas", new byte[] { 65 });

            Commit();

            Assert.IsFalse(File.Exists(marker));
        }

        [TestCase("../outside.bas")]
        [TestCase("-rf")]
        [TestCase("C:/Windows/win.ini")]
        [TestCase(".git/config")]
        [TestCase("")]
        public void ReadFileAt_RefusesAPathThatIsNotInTheWorkTree(string path)
        {
            _repository.EnsureCreated();
            Write("a.bas", new byte[] { 65 });
            var id = Commit();

            Assert.Throws<ArgumentException>(() => _repository.ReadFileAt(id, path));
        }

        [TestCase("--output=x")]
        [TestCase("HEAD")]
        [TestCase("zzzzzzz")]
        public void ReadFileAt_RefusesAnythingButACommitId(string commit)
        {
            _repository.EnsureCreated();

            Assert.Throws<ArgumentException>(() => _repository.ReadFileAt(commit, "a.bas"));
        }

        [Test]
        public void Run_FailsWithWhatGitSaid()
        {
            var exception = Assert.Throws<GitException>(() => _git.RunOrThrow(new[] { "--git-dir=" + Path.Combine(_folder, "none"), "log" }));

            StringAssert.Contains("failed", exception.Message);
        }

        [Test]
        public void Run_MasksSecretsInWhatItReports()
        {
            var runner = new GitRunner(_gitPath, secrets: new[] { "s3cr3t-token" });

            var exception = Assert.Throws<GitException>(() => runner.RunOrThrow(new[] { "--no-such-option=s3cr3t-token" }));

            StringAssert.DoesNotContain("s3cr3t-token", exception.Message);
        }

        [TestCase("plain", "plain")]
        [TestCase("with space", "\"with space\"")]
        [TestCase("", "\"\"")]
        [TestCase("say \"hi\"", "\"say \\\"hi\\\"\"")]
        [TestCase("C:\\dir with space\\", "\"C:\\dir with space\\\\\"")]
        [TestCase("a\\\"b", "\"a\\\\\\\"b\"")]
        public void Quote_FollowsTheWindowsCommandLineRules(string argument, string expected)
            => Assert.AreEqual(expected, GitRunner.Quote(argument));
    }
}
