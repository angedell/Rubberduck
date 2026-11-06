using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Abstractions;
using System.Linq;
using System.Text;
using NUnit.Framework;
using Rubberduck.Git;

namespace RubberduckTests.Git
{
    /// <summary>Real git against a bare repository in a temp folder standing in for GitHub; no network, no credentials.</summary>
    [TestFixture]
    public class GitRemoteTests
    {
        private string _folder;
        private string _remote;
        private string _other;
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

            _folder = Path.Combine(Path.GetTempPath(), "rd-remote-" + Guid.NewGuid());
            _remote = Path.Combine(_folder, "remote.git");
            _other = Path.Combine(_folder, "other");
            Directory.CreateDirectory(_folder);
            _git = new GitRunner(installation.Path);
            _git.RunOrThrow(new[] { "init", "--bare", "--quiet", "--initial-branch=main", _remote });

            _repository = new HiddenRepository(_git, new FileSystem(), Path.Combine(_folder, "mine"));
            _repository.EnsureCreated();
            _repository.SetRemote(_remote);
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

        private void Mine(string name, string content)
        {
            File.WriteAllText(Path.Combine(_repository.WorkTree, name), content, Encoding.ASCII);
        }

        private string MineCommit(string message = "mine") => _repository.Commit(message, "Me", "me@example.test");

        //Somebody else: a plain clone with its own work tree, pushing to the same remote.
        private void Other(string name, string content, string message)
        {
            if (!Directory.Exists(_other))
            {
                _git.RunOrThrow(new[] { "clone", "--quiet", _remote, _other });
                //A plain git on Windows converts line endings by default; the hidden repository never does, and this stands in for a collaborator who does not either.
                _git.RunOrThrow(new[] { "-C", _other, "config", "core.autocrlf", "false" });
            }
            else
            {
                _git.RunOrThrow(new[] { "-C", _other, "pull", "--quiet", "--no-rebase" });
            }

            File.WriteAllText(Path.Combine(_other, name), content, Encoding.ASCII);
            _git.RunOrThrow(new[] { "-C", _other, "add", "--all" });
            _git.RunOrThrow(new[] { "-C", _other, "-c", "user.name=Other", "-c", "user.email=o@example.test", "commit", "--quiet", "--message", message });
            _git.RunOrThrow(new[] { "-C", _other, "push", "--quiet", "origin", "HEAD:main" });
        }

        [Test]
        public void Push_SendsTheBranchAndSetsItsUpstream()
        {
            Mine("a.bas", "one\r\n");
            var id = MineCommit();

            _repository.Push();
            _repository.Fetch();

            Assert.AreEqual(id, _git.RunOrThrow(new[] { "--git-dir=" + _remote, "rev-parse", "main" }).Output.Trim());
            Assert.AreEqual((0, 0), _repository.AheadBehind());
        }

        [Test]
        public void AheadBehind_IsNullForARemoteThatHasNoSuchBranchYet()
        {
            Mine("a.bas", "one\r\n");
            MineCommit();
            _repository.Fetch();

            Assert.IsNull(_repository.AheadBehind());
        }

        [Test]
        public void Push_OfALocalCommitMakesTheBranchAhead_AndFetchShowsWhatTheOtherSideAdded()
        {
            Mine("a.bas", "one\r\n");
            MineCommit();
            _repository.Push();
            Mine("a.bas", "two\r\n");
            MineCommit("second");
            Other("b.bas", "theirs\r\n", "theirs");
            _repository.Fetch();

            Assert.AreEqual((1, 1), _repository.AheadBehind());
        }

        [Test]
        public void MergeRemote_FastForwardsWhenThisSideHasNothingOfItsOwn()
        {
            Mine("a.bas", "one\r\n");
            MineCommit();
            _repository.Push();
            Other("b.bas", "theirs\r\n", "theirs");
            _repository.Fetch();

            var outcome = _repository.MergeRemote("Me", "me@example.test");

            Assert.AreEqual(GitMergeKind.FastForwarded, outcome.Kind);
            Assert.AreEqual(new[] { "b.bas" }, outcome.Changes.Select(change => change.Path).ToArray());
            Assert.AreEqual(GitChangeKind.Added, outcome.Changes[0].Kind);
            Assert.AreEqual("theirs\r\n", File.ReadAllText(Path.Combine(_repository.WorkTree, "b.bas")));
        }

        [Test]
        public void MergeRemote_MergesChangesToDifferentFiles()
        {
            Mine("a.bas", "one\r\n");
            MineCommit();
            _repository.Push();
            Mine("c.bas", "mine\r\n");
            MineCommit("mine");
            Other("b.bas", "theirs\r\n", "theirs");
            _repository.Fetch();

            var outcome = _repository.MergeRemote("Me", "me@example.test");

            Assert.AreEqual(GitMergeKind.Merged, outcome.Kind);
            Assert.AreEqual(new[] { "b.bas" }, outcome.Changes.Select(change => change.Path).ToArray());
            Assert.IsTrue(File.Exists(Path.Combine(_repository.WorkTree, "c.bas")));
            Assert.AreEqual((2, 0), _repository.AheadBehind());
        }

        [Test]
        public void MergeRemote_WhenTheSameLinesChanged_UndoesTheMergeAndNamesTheFile()
        {
            Mine("a.bas", "one\r\n");
            var first = MineCommit();
            _repository.Push();
            Mine("a.bas", "mine\r\n");
            var mine = MineCommit("mine");
            Other("a.bas", "theirs\r\n", "theirs");
            _repository.Fetch();

            var outcome = _repository.MergeRemote("Me", "me@example.test");

            Assert.AreEqual(GitMergeKind.Conflicts, outcome.Kind);
            CollectionAssert.AreEqual(new[] { "a.bas" }, outcome.Conflicts);
            Assert.AreEqual("mine\r\n", File.ReadAllText(Path.Combine(_repository.WorkTree, "a.bas")));
            Assert.IsEmpty(_repository.Status());
            Assert.AreEqual(mine, _repository.Log(1)[0].Id);
            Assert.AreNotEqual(first, mine);
        }

        [Test]
        public void MergeRemote_WithNothingNewOrNothingOnTheRemote_IsUpToDate()
        {
            Mine("a.bas", "one\r\n");
            MineCommit();
            _repository.Fetch();
            Assert.AreEqual(GitMergeKind.UpToDate, _repository.MergeRemote("Me", "me@example.test").Kind);

            _repository.Push();
            _repository.Fetch();
            Assert.AreEqual(GitMergeKind.UpToDate, _repository.MergeRemote("Me", "me@example.test").Kind);
        }

        [Test]
        public void Push_IsRefusedWhenTheRemoteHasMovedOn_AndNothingIsForced()
        {
            Mine("a.bas", "one\r\n");
            MineCommit();
            _repository.Push();
            Other("b.bas", "theirs\r\n", "theirs");
            Mine("a.bas", "two\r\n");
            MineCommit("second");

            Assert.Throws<GitException>(() => _repository.Push());
            Assert.AreEqual("theirs", _git.RunOrThrow(new[] { "--git-dir=" + _remote, "log", "-1", "--format=%s", "main" }).Output.Trim());
        }

        [Test]
        public void SetRemote_AddsThenChangesTheUrl()
        {
            Assert.AreEqual(_remote, _repository.RemoteUrl());

            _repository.SetRemote("https://example.test/o/r.git");

            Assert.AreEqual("https://example.test/o/r.git", _repository.RemoteUrl());
        }

        // ---- GitRemoteAuth ----

        [TestCase("https://github.com/owner/repo.git", "github.com")]
        [TestCase("https://GitHub.com/owner/repo", "github.com")]
        [TestCase("https://dev.azure.com/org/proj/_git/repo", "dev.azure.com")]
        public void HostOf_AcceptsHttpsUrls(string url, string host) => Assert.AreEqual(host, GitRemoteAuth.HostOf(url));

        [TestCase("git@github.com:owner/repo.git")]
        [TestCase("ssh://git@github.com/owner/repo.git")]
        [TestCase("http://github.com/owner/repo.git")]
        [TestCase("ext::sh -c touch% /tmp/pwned")]
        [TestCase(@"C:\repos\thing.git")]
        [TestCase("file:///C:/repos/thing.git")]
        [TestCase("https://user:token@github.com/owner/repo.git")]
        [TestCase("")]
        [TestCase(null)]
        public void HostOf_RefusesEverythingElse(string url) => Assert.IsNull(GitRemoteAuth.HostOf(url));

        [Test]
        public void ForHost_PassesTheTokenAsAHeaderForThatHostOnly_AndMasksIt()
        {
            var auth = GitRemoteAuth.ForHost("github.com", new StoredCredential("angelo", "ghp_secret"));

            Assert.AreEqual("1", auth.Environment["GIT_CONFIG_COUNT"]);
            Assert.AreEqual("http.https://github.com/.extraHeader", auth.Environment["GIT_CONFIG_KEY_0"]);
            var header = auth.Environment["GIT_CONFIG_VALUE_0"];
            StringAssert.StartsWith("Authorization: Basic ", header);
            Assert.AreEqual("angelo:ghp_secret", Encoding.UTF8.GetString(Convert.FromBase64String(header.Substring("Authorization: Basic ".Length))));
            CollectionAssert.Contains(auth.Secrets, "ghp_secret");
            CollectionAssert.Contains(auth.Secrets, header.Substring("Authorization: Basic ".Length));
        }

        [Test]
        public void ForHost_WithoutAUserName_UsesAPlaceholderGitHubAccepts()
        {
            var header = GitRemoteAuth.ForHost("github.com", new StoredCredential(null, "t")).Environment["GIT_CONFIG_VALUE_0"];

            Assert.AreEqual("x-access-token:t", Encoding.UTF8.GetString(Convert.FromBase64String(header.Substring("Authorization: Basic ".Length))));
        }

        // ---- Credential Manager (the real one, under a throwaway target) ----

        [Test]
        public void CredentialStore_KeepsAndForgetsASecret()
        {
            var store = new WindowsCredentialStore();
            var target = CredentialTargets.Prefix + "test/" + Guid.NewGuid();
            try
            {
                Assert.IsNull(store.Read(target));

                store.Write(target, "angelo", "s3cr3t \u00e9 with spaces");
                var read = store.Read(target);

                Assert.AreEqual("angelo", read.User);
                Assert.AreEqual("s3cr3t \u00e9 with spaces", read.Secret);
                CollectionAssert.Contains(store.List(CredentialTargets.Prefix + "test/"), target);

                store.Write(target, "angelo", "changed");
                Assert.AreEqual("changed", store.Read(target).Secret);
                Assert.IsTrue(store.Delete(target));
                Assert.IsNull(store.Read(target));
                Assert.IsFalse(store.Delete(target));
            }
            finally
            {
                store.Delete(target);
            }
        }

        [Test]
        public void CredentialTargets_AreNamedByHostAndProvider()
        {
            Assert.AreEqual("Rubberduck/git/github.com", CredentialTargets.GitHost("GitHub.com"));
            Assert.AreEqual("Rubberduck/ai/anthropic", CredentialTargets.Ai("Anthropic"));
        }
    }
}
