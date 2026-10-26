using System;
using NUnit.Framework;
using Rubberduck.UI.CodeExplorer.GitSync;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    [Category("GitSync")]
    public class GitSyncStampTests
    {
        private const string ClassFile =
            "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Foo\"\r\nAttribute VB_Exposed = False\r\nOption Explicit\r\nPublic X As Long\r\n";

        [Test]
        public void Format_AndParse_RoundTrip()
        {
            var stamp = GitSyncStamp.Parse("Option Explicit\r\n" + GitSyncStamp.Format(12, "3fa9c1d4e07b22aa") + "\r\n");

            Assert.AreEqual(12, stamp.Revision);
            Assert.AreEqual("3fa9c1d4e07b22aa", stamp.Hash);
        }

        [Test]
        public void Format_WithTime_RoundTrips()
        {
            var utc = new DateTime(2025, 6, 1, 12, 34, 56, DateTimeKind.Utc);
            var line = GitSyncStamp.Format(7, "abcdef0123456789", utc);

            Assert.AreEqual("' rdsync rev=7 hash=abcdef0123456789 utc=2025-06-01T12:34:56Z", line);

            var stamp = GitSyncStamp.Parse(line);
            Assert.AreEqual(utc, stamp.Utc);
            Assert.AreEqual(DateTimeKind.Utc, stamp.Utc.Value.Kind);
        }

        [Test]
        public void Parse_StampWithoutTime_HasNoTime()
        {
            Assert.IsNull(GitSyncStamp.Parse("' rdsync rev=7 hash=abcdef").Utc);
        }

        [Test]
        public void StampWithTimeDoesNotChangeTheCodeHash()
        {
            var plain = GitSyncStamp.HashOfCode("Option Explicit");
            var stamped = GitSyncStamp.HashOfCode("' rdsync rev=2 hash=ab utc=2025-06-01T12:34:56Z\r\nOption Explicit");

            Assert.AreEqual(plain, stamped);
        }

        [Test]
        public void Parse_WithoutStamp_ReturnsNull()
        {
            Assert.IsNull(GitSyncStamp.Parse("Option Explicit\r\n' just a comment\r\n"));
        }

        [Test]
        public void HashOfCode_IgnoresStampLineLineEndingsAndOuterBlankSpace()
        {
            var plain = GitSyncStamp.HashOfCode("Option Explicit\r\nPublic X As Long");
            var stamped = GitSyncStamp.HashOfCode("' rdsync rev=3 hash=abcdef\nOption Explicit\nPublic X As Long\n\n");

            Assert.AreEqual(plain, stamped);
        }

        [Test]
        public void HashOfCode_ChangesWhenCodeChanges()
        {
            Assert.AreNotEqual(GitSyncStamp.HashOfCode("Public X As Long"), GitSyncStamp.HashOfCode("Public Y As Long"));
        }

        [Test]
        public void CodeOfFile_DropsTheHeader()
        {
            Assert.AreEqual("Option Explicit\r\nPublic X As Long\r\n", GitSyncStamp.CodeOfFile(ClassFile));
        }

        [Test]
        public void CodeOfFile_DropsNestedFormDesignerBlock()
        {
            var form = "VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 \r\n   Begin Foo.Button Button1 \r\n      Caption = \"x\"\r\n   End\r\nEnd\r\nAttribute VB_Name = \"UserForm1\"\r\nOption Explicit\r\n";

            Assert.AreEqual("Option Explicit\r\n", GitSyncStamp.CodeOfFile(form));
        }

        [Test]
        public void CodeOfFile_WithoutHeader_IsTheWholeText()
        {
            Assert.AreEqual("Option Explicit", GitSyncStamp.CodeOfFile("Option Explicit"));
        }

        [Test]
        public void WithStamp_InsertsAfterAttributes_AndReplacesOnSecondCall()
        {
            var first = GitSyncStamp.WithStamp(ClassFile, 1, "aaaa");
            StringAssert.Contains("Attribute VB_Exposed = False\r\n' rdsync rev=1 hash=aaaa\r\nOption Explicit", first);

            var second = GitSyncStamp.WithStamp(first, 2, "bbbb");
            StringAssert.Contains("rev=2 hash=bbbb", second);
            StringAssert.DoesNotContain("rev=1", second);
        }

        [Test]
        public void StampedFileMatchesTheModuleCodeAfterImport()
        {
            //What the VBE keeps in the module after importing the stamped file: the file without its header.
            var stamped = GitSyncStamp.WithStamp(ClassFile, 4, GitSyncStamp.HashOfCode(GitSyncStamp.CodeOfFile(ClassFile)));
            var moduleCode = GitSyncStamp.CodeOfFile(stamped);

            var file = GitSyncStamp.StateOfFile(stamped);
            var module = GitSyncStamp.StateOfCode(moduleCode);

            Assert.IsFalse(file.Dirty);
            Assert.IsFalse(module.Dirty);
            Assert.AreEqual(file.BodyHash, module.BodyHash);
            Assert.AreEqual(4, module.Revision);
        }

        [Test]
        public void EditedCodeMakesTheStateDirty()
        {
            var stamped = GitSyncStamp.WithStamp(ClassFile, 4, GitSyncStamp.HashOfCode(GitSyncStamp.CodeOfFile(ClassFile)));

            Assert.IsTrue(GitSyncStamp.StateOfFile(stamped.Replace("Long", "String")).Dirty);
        }
    }

    [TestFixture]
    [Category("GitSync")]
    public class StampDirectionTests
    {
        private static readonly DateTime Now = new DateTime(2025, 1, 1, 12, 0, 0, DateTimeKind.Utc);

        private static StampState Clean(int revision, string hash = "h")
        {
            return new StampState(new SyncStamp(revision, hash), hash);
        }

        private static StampState Dirty(int revision, string bodyHash = "edited")
        {
            return new StampState(new SyncStamp(revision, "h"), bodyHash);
        }

        private static SyncDecision Decide(StampState module, StampState file, DateTime? fileWritten = null, DateTime? moduleChanged = null)
        {
            return SyncDirectionResolver.DecideByStamps(module, file, fileWritten ?? Now, moduleChanged ?? Now);
        }

        [Test]
        public void StaleModuleOnAnotherMachine_ImportsTheNewerFile()
        {
            //Machine A pushed revision 5; the workbook on machine B is still at revision 3 and was not edited.
            var decision = Decide(Clean(3, "old"), Clean(5, "new"));

            Assert.AreEqual(SyncDirection.Import, decision.Direction);
            Assert.IsFalse(decision.IsConflict);
        }

        [Test]
        public void NewerCleanModule_ExportsToTheStaleFile()
        {
            Assert.AreEqual(SyncDirection.Export, Decide(Clean(5, "new"), Clean(3, "old")).Direction);
        }

        [Test]
        public void SameRevisionAndHash_DoesNothing()
        {
            Assert.AreEqual(SyncDirection.None, Decide(Clean(4), Clean(4)).Direction);
        }

        [Test]
        public void SameRevisionButDifferentContent_IsAConflict()
        {
            Assert.IsTrue(Decide(Clean(4, "a"), Clean(4, "b")).IsConflict);
        }

        private static StampState CleanAt(int revision, string hash, DateTime utc)
        {
            return new StampState(new SyncStamp(revision, hash, utc), hash);
        }

        [Test]
        public void SameRevisionButDifferentContent_TheStampTimesDecide()
        {
            var older = Now.AddHours(-3);

            //The clocks of this machine say the file was just written, but the stamps know better.
            var moduleNewer = Decide(CleanAt(4, "a", Now), CleanAt(4, "b", older), fileWritten: Now, moduleChanged: older);
            var fileNewer = Decide(CleanAt(4, "a", older), CleanAt(4, "b", Now), fileWritten: older, moduleChanged: Now);

            Assert.AreEqual(SyncDirection.Export, moduleNewer.Direction);
            Assert.AreEqual(SyncDirection.Import, fileNewer.Direction);
        }

        [Test]
        public void EditedModule_ExportsOverACleanFileOfTheSameRevision()
        {
            var decision = Decide(Dirty(4), Clean(4));

            Assert.AreEqual(SyncDirection.Export, decision.Direction);
            Assert.IsFalse(decision.IsConflict);
        }

        [Test]
        public void EditedModuleButNewerFile_IsAConflict()
        {
            Assert.IsTrue(Decide(Dirty(3), Clean(5)).IsConflict);
        }

        [Test]
        public void EditedFile_ImportsOverACleanModuleOfTheSameRevision()
        {
            var decision = Decide(Clean(4), Dirty(4));

            Assert.AreEqual(SyncDirection.Import, decision.Direction);
            Assert.IsFalse(decision.IsConflict);
        }

        [Test]
        public void EditedFileButNewerModule_IsAConflict()
        {
            Assert.IsTrue(Decide(Clean(6), Dirty(4)).IsConflict);
        }

        [Test]
        public void BothEditedToTheSameCode_OnlyRecords()
        {
            Assert.AreEqual(SyncDirection.Record, Decide(Dirty(4, "same"), Dirty(4, "same")).Direction);
        }

        [Test]
        public void BothEditedDifferently_NewerSideWins()
        {
            var fileNewer = Decide(Dirty(4, "a"), Dirty(4, "b"), fileWritten: Now, moduleChanged: Now.AddMinutes(-5));
            var moduleNewer = Decide(Dirty(4, "a"), Dirty(4, "b"), fileWritten: Now.AddMinutes(-5), moduleChanged: Now);

            Assert.AreEqual(SyncDirection.Import, fileNewer.Direction);
            Assert.AreEqual(SyncDirection.Export, moduleNewer.Direction);
            Assert.IsTrue(fileNewer.IsConflict && moduleNewer.IsConflict);
        }
    }
}
