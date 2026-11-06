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
    /// <summary>Real git: whether a patch cut into hunks applies, and lands, exactly as the text the VBE wrote.</summary>
    [TestFixture]
    public class GitSplitTests
    {
        private string _folder;
        private HiddenRepository _repository;
        private static readonly Encoding Ansi = GitHunkParser.Bytes;

        [SetUp]
        public void SetUp()
        {
            var installation = new GitLocator(new FileSystem(), GitLocator.RunVersion, null, Environment.GetEnvironmentVariable("PATH"), GitLocator.DefaultOwnFolder,
                Environment.GetEnvironmentVariable("ProgramFiles")).Find();
            if (installation == null)
            {
                Assert.Ignore("No git on this machine.");
            }

            _folder = Path.Combine(Path.GetTempPath(), "rd-split-" + Guid.NewGuid());
            _repository = new HiddenRepository(new GitRunner(installation.Path), new FileSystem(), _folder);
            _repository.EnsureCreated();
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

        private void Write(string name, byte[] bytes) => File.WriteAllBytes(Path.Combine(_repository.WorkTree, name), bytes);

        private byte[] Read(string name) => File.ReadAllBytes(Path.Combine(_repository.WorkTree, name));

        //Forty numbered lines, CRLF, one of them with a letter that is a single byte in the ANSI code page.
        private static byte[] Lines(string third = "line 3", string thirtyFifth = "line 35")
            => Ansi.GetBytes(string.Concat(Enumerable.Range(1, 40).Select(i =>
                (i == 3 ? third : i == 35 ? thirtyFifth : i == 20 ? "café 20" : "line " + i) + "\r\n")));

        private string Commit(string message = "base") => _repository.Commit(message, "T", "t@example.test");

        private IReadOnlyList<GitHunk> Hunks() => GitHunkParser.Parse(_repository.StagedPatch());

        [Test]
        public void Parse_CutsAFileIntoItsBlocksAndKeepsBinaryFilesWhole()
        {
            Write("Module1.bas", Lines());
            Write("Form1.frx", new byte[] { 0, 1, 2, 3, 255 });
            Commit();
            Write("Module1.bas", Lines("changed 3", "changed 35"));
            Write("Form1.frx", new byte[] { 0, 1, 2, 3, 254 });

            var hunks = Hunks();

            CollectionAssert.AreEqual(new[] { "Form1.frx", "Module1.bas", "Module1.bas" }, hunks.Select(hunk => hunk.Path).ToArray());
            CollectionAssert.AreEqual(new[] { "h1", "h2", "h3" }, hunks.Select(hunk => hunk.Id).ToArray());
            Assert.IsTrue(hunks[0].IsWholeFile && hunks[0].IsBinary && hunks[0].IsCompanion);
            StringAssert.Contains("changed 3", hunks[1].Body);
            StringAssert.Contains("changed 35", hunks[2].Body);
            StringAssert.DoesNotContain("changed 35", hunks[1].Body);
        }

        [Test]
        public void CommitParts_MakesOneCommitPerPart_AndTheLastOneEqualsTheWorkTreeByteForByte()
        {
            Write("Module1.bas", Lines());
            Commit();
            var changed = Lines("changed 3", "changed 35");
            Write("Module1.bas", changed);
            var hunks = Hunks();

            var ids = _repository.CommitParts(new[]
            {
                new GitCommitPart("Change the start", new[] { hunks[0] }),
                new GitCommitPart("Change the end", new[] { hunks[1] })
            }, "T", "t@example.test");

            Assert.AreEqual(2, ids.Count);
            Assert.AreEqual(new[] { "Change the end", "Change the start", "base" }, _repository.Log().Select(commit => commit.Subject).ToArray());
            CollectionAssert.AreEqual(Lines("changed 3", "line 35"), _repository.ReadFileAt(ids[0], "Module1.bas"));
            CollectionAssert.AreEqual(changed, _repository.ReadFileAt(ids[1], "Module1.bas"));
            CollectionAssert.AreEqual(changed, Read("Module1.bas"));
            Assert.IsEmpty(_repository.Status());
        }

        [Test]
        public void CommitParts_CanSplitTheVeryFirstCommit()
        {
            Write("A.bas", Ansi.GetBytes("a\r\n"));
            Write("B.bas", Ansi.GetBytes("b\r\n"));
            var hunks = Hunks();

            _repository.CommitParts(new[]
            {
                new GitCommitPart("Add B", hunks.Where(hunk => hunk.Path == "B.bas").ToList()),
                new GitCommitPart("Add A", hunks.Where(hunk => hunk.Path == "A.bas").ToList())
            }, "T", "t@example.test");

            Assert.AreEqual(new[] { "Add A", "Add B" }, _repository.Log().Select(commit => commit.Subject).ToArray());
            Assert.IsEmpty(_repository.Status());
        }

        [Test]
        public void CommitParts_PutsEverythingBackWhenAPartDoesNotApply()
        {
            Write("Module1.bas", Lines());
            Commit();
            Write("Module1.bas", Lines("changed 3", "changed 35"));
            var hunks = Hunks();

            Assert.Throws<GitException>(() => _repository.CommitParts(new[]
            {
                new GitCommitPart("First", new[] { hunks[0] }),
                new GitCommitPart("Again", new[] { hunks[0] })
            }, "T", "t@example.test"));

            Assert.AreEqual(new[] { "base" }, _repository.Log().Select(commit => commit.Subject).ToArray());
            CollectionAssert.AreEqual(Lines("changed 3", "changed 35"), Read("Module1.bas"));
            Assert.AreEqual(1, _repository.Status().Count);
        }

        [Test]
        public void CommitParts_RefusesPartsThatLeaveAChangeOut_AndPutsEverythingBack()
        {
            Write("Module1.bas", Lines());
            Commit();
            Write("Module1.bas", Lines("changed 3", "changed 35"));
            var hunks = Hunks();

            Assert.Throws<GitException>(() => _repository.CommitParts(new[] { new GitCommitPart("Half", new[] { hunks[0] }) }, "T", "t@example.test"));

            Assert.AreEqual(new[] { "base" }, _repository.Log().Select(commit => commit.Subject).ToArray());
            CollectionAssert.AreEqual(Lines("changed 3", "changed 35"), Read("Module1.bas"));
        }

        [Test]
        public void CommitParts_RefusesAnEmptyPartOrMessage()
        {
            Write("Module1.bas", Lines());
            var hunks = Hunks();

            Assert.Throws<ArgumentException>(() => _repository.CommitParts(new[] { new GitCommitPart(" ", hunks) }, "T", "t@example.test"));
            Assert.Throws<ArgumentException>(() => _repository.CommitParts(new[] { new GitCommitPart("x", new GitHunk[0]) }, "T", "t@example.test"));
        }
    }
}
