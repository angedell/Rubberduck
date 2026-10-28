using System.Collections.Generic;
using System.IO;
using System.IO.Abstractions;
using System.Linq;
using NUnit.Framework;
using Rubberduck.UI.CodeExplorer.GitSync;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    public class GitSyncNotesTests
    {
        private string _root;
        private GitSyncNotes _sync;

        [SetUp]
        public void SetUp()
        {
            _root = Path.Combine(Path.GetTempPath(), "rd-notes-" + System.Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_root);
            var fileSystem = new FileSystem();
            _sync = new GitSyncNotes(fileSystem, new GitSyncFolderProcessor(fileSystem));
        }

        [TearDown]
        public void TearDown() => Directory.Delete(_root, true);

        private static NoteInput Note(string module, string name, string folder, string markdown)
        {
            var lines = new List<string> { "Attribute VB_Name = \"" + module + "\"", "'@Notes(\"" + name + "\")" };
            if (folder != null)
            {
                lines.Add("'@Folder(\"" + folder + "\")");
            }

            lines.AddRange(markdown.Split('\n').Select(line => line.Length == 0 ? "'" : "' " + line));
            return new NoteInput { ModuleName = module, Code = string.Join("\r\n", lines) };
        }

        private SyncPlan Plan(params NoteInput[] notes)
        {
            var plan = new SyncPlan();
            _sync.Plan(notes, _root, new GitSyncState(), new GitIgnoreSet(new List<KeyValuePair<string, GitIgnore>>()), plan);
            return plan;
        }

        private void WriteFile(string relative, string text)
        {
            var path = Path.Combine(_root, relative);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, text);
        }

        [Test]
        public void ANoteWithoutAFile_IsExportedToItsFolder_AsMarkdown()
        {
            var plan = Plan(Note("CLAUDE_md", "CLAUDE.md", "Api.Core", "# Title"));

            var item = plan.Items.Single();
            Assert.AreEqual("Api/Core/CLAUDE.md", item.Name);
            Assert.AreEqual(@"Api\Core\CLAUDE.md", item.RelativePath);
            Assert.AreEqual("CLAUDE_md", item.NoteModule);
            Assert.AreEqual(SyncStatus.ModuleOnly, item.Status);
            Assert.AreEqual(SyncAction.Export, item.Suggested);
            Assert.IsTrue(item.IsNote);
        }

        [Test]
        public void ANoteWithoutAFolder_GoesToTheRoot()
        {
            Assert.AreEqual("CLAUDE.md", Plan(Note("CLAUDE_md", "CLAUDE.md", null, "x")).Items.Single().RelativePath);
        }

        [Test]
        public void TheSameText_IsInSync_WhateverTheLineEndingsAndTrailingBlankLines()
        {
            WriteFile("CLAUDE.md", "﻿# Title\r\n\r\ntext\r\n\r\n");
            var plan = Plan(Note("CLAUDE_md", "CLAUDE.md", null, "# Title\n\ntext"));

            Assert.AreEqual(0, plan.Items.Count);
            Assert.AreEqual(1, plan.InSyncCount);
            Assert.IsTrue(plan.InSyncBaselines.ContainsKey("CLAUDE.md"));
        }

        [Test]
        public void AFileOnlyOfTheFolder_IsListedButNeverSuggestedForImport()
        {
            WriteFile("README.md", "# Readme");
            WriteFile(@"docs\Guide.md", "guide");
            WriteFile(@".git\notes.md", "hidden");

            var plan = Plan();

            CollectionAssert.AreEquivalent(new[] { "README.md", "docs/Guide.md" }, plan.Items.Select(i => i.Name));
            Assert.IsTrue(plan.Items.All(i => i.Status == SyncStatus.FileOnly && i.Suggested == SyncAction.None && !i.ModuleExists));
        }

        [Test]
        public void ADifferingNote_CanBeCompared()
        {
            WriteFile("CLAUDE.md", "old");
            var item = Plan(Note("CLAUDE_md", "CLAUDE.md", null, "new")).Items.Single();

            Assert.AreEqual(SyncStatus.Differs, item.Status);
            Assert.AreEqual("new", item.ModuleText);
            Assert.AreEqual("old", item.FileText);
            Assert.IsTrue(item.CanCompare);
        }

        [Test]
        public void ANoteThatMovedFolder_IsAnotherPath()
        {
            WriteFile(@"Old\CLAUDE.md", "same");
            var plan = Plan(Note("CLAUDE_md", "CLAUDE.md", "New", "same"));

            CollectionAssert.AreEquivalent(new[] { "New/CLAUDE.md", "Old/CLAUDE.md" }, plan.Items.Select(i => i.Name));
        }
    }
}
