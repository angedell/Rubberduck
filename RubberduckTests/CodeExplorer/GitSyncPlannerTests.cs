using System.Linq;
using NUnit.Framework;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.CodeExplorer.GitSync;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    [Category("GitSync")]
    public class SyncTextTests
    {
        private static readonly IIndenterSettings Settings = new IndenterSettings(true);
        private static readonly ISimpleIndenter Indenter = new Indenter(null, () => Settings);

        private const string Header = "Attribute VB_Name = \"Foo\"\r\n";

        private static string Signature(string text) => SyncText.Signature(text, Settings, Indenter);

        [Test]
        public void Signature_IgnoresLineEndings()
        {
            var crlf = Header + "Option Explicit\r\nPublic Sub A()\r\nEnd Sub\r\n";

            Assert.AreEqual(Signature(crlf), Signature(crlf.Replace("\r\n", "\n")));
        }

        [Test]
        public void Signature_IgnoresIndentation()
        {
            var flat = Header + "Public Sub A()\r\nDebug.Print 1\r\nEnd Sub\r\n";
            var indented = Header + "Public Sub A()\r\n        Debug.Print 1\r\nEnd Sub\r\n";

            Assert.AreEqual(Signature(flat), Signature(indented));
        }

        [Test]
        public void Signature_IgnoresLegacyStampsAndTrailingSpace()
        {
            var plain = Header + "Option Explicit\r\n";
            var stamped = Header + "' rdsync rev=3 hash=3fa9c1d4e07b22aa utc=2025-06-01T12:34:56Z\r\nOption Explicit   \r\n";

            Assert.AreEqual(Signature(plain), Signature(stamped));
        }

        [Test]
        public void Signature_NoticesRealCodeChanges()
        {
            Assert.AreNotEqual(Signature(Header + "Public Sub A()\r\nEnd Sub\r\n"), Signature(Header + "Public Sub B()\r\nEnd Sub\r\n"));
        }

        [Test]
        public void Signature_WithCaseIgnored_IgnoresCaseOutsideStrings()
        {
            var lower = Header + "Public Sub a()\r\nDim x As Long ' note\r\nEnd Sub\r\n";
            var upper = Header + "PUBLIC SUB A()\r\ndim X as long ' NOTE\r\nEND SUB\r\n";

            Assert.AreEqual(SyncText.Signature(lower, Settings, Indenter, true), SyncText.Signature(upper, Settings, Indenter, true));
            Assert.AreNotEqual(Signature(lower), Signature(upper), "the case counts unless asked otherwise");
        }

        [Test]
        public void Signature_WithCaseIgnored_StillSeesCaseInsideStrings()
        {
            var one = Header + "Public Sub A()\r\nDebug.Print \"Hello\"\r\nEnd Sub\r\n";
            var other = Header + "Public Sub A()\r\nDebug.Print \"HELLO\"\r\nEnd Sub\r\n";
            var quoted = Header + "Public Sub A()\r\nDebug.Print \"a \"\"B\"\" c\"\r\nEnd Sub\r\n";
            var quotedOther = Header + "Public Sub A()\r\nDebug.Print \"a \"\"b\"\" c\"\r\nEnd Sub\r\n";

            Assert.AreNotEqual(SyncText.Signature(one, Settings, Indenter, true), SyncText.Signature(other, Settings, Indenter, true));
            Assert.AreNotEqual(SyncText.Signature(quoted, Settings, Indenter, true), SyncText.Signature(quotedOther, Settings, Indenter, true), "a doubled quote does not end the string");
        }

        [Test]
        public void FoldCase_KeepsLengthAndStrings()
        {
            const string line = "Dim S As String: S = \"AbC\" & \"x\"\"Y\" ' ToDo";

            Assert.AreEqual("dim s as string: s = \"AbC\" & \"x\"\"Y\" ' todo", SyncText.FoldCase(line));
        }

        [Test]
        public void Signature_NoticesAttributeChanges()
        {
            var plain = Header + "Option Explicit\r\n";
            var predeclared = Header + "Attribute VB_PredeclaredId = True\r\nOption Explicit\r\n";

            Assert.AreNotEqual(Signature(plain), Signature(predeclared));
        }

        [Test]
        public void Signature_LeavesNoIndentModulesAsTheyAre()
        {
            var first = Header + "'@NoIndent\r\nPublic Sub A()\r\nDebug.Print 1\r\nEnd Sub\r\n";
            var second = Header + "'@NoIndent\r\nPublic Sub A()\r\n    Debug.Print 1\r\nEnd Sub\r\n";

            Assert.AreNotEqual(Signature(first), Signature(second));
        }

        [Test]
        public void PrepareForImport_FixesLineEndingsAndAnnotation()
        {
            var prepared = SyncText.PrepareForImport("Attribute VB_Name = \"Foo\"\nOption Explicit\n", "Models\\Orders");

            Assert.AreEqual("Attribute VB_Name = \"Foo\"\r\n'@Folder(\"Models.Orders\")\r\nOption Explicit\r\n", prepared);
        }

        [Test]
        public void PrepareForImport_LeavesRootFilesAnnotationAlone()
        {
            var text = "Attribute VB_Name = \"Foo\"\r\n'@Folder(\"Elsewhere\")\r\nOption Explicit\r\n";

            Assert.AreEqual(text, SyncText.PrepareForImport(text, string.Empty));
        }

        [Test]
        public void CodeStart_SkipsFormHeaderAndAttributes()
        {
            var lines = SyncText.Lines("VERSION 5.00\r\nBegin {X} Foo\r\n   Caption = \"a\"\r\nEnd\r\nAttribute VB_Name = \"Foo\"\r\nOption Explicit");

            Assert.AreEqual(5, SyncText.CodeStart(lines));
        }
    }

    [TestFixture]
    [Category("GitSync")]
    public class SyncHygieneTests
    {
        private static SyncConfig Enforcing(params string[] rules)
        {
            var config = new SyncConfig { OnViolation = ViolationMode.Block };
            config.Enforce.AddRange(rules);
            return config;
        }

        private static string[] Rules(string text, SyncConfig config, bool folderFollowsPath = false) =>
            SyncHygiene.Check("Foo", text, SyncSide.Module, config, folderFollowsPath).Select(v => v.Rule).ToArray();

        [Test]
        public void NothingIsChecked_IfNothingIsEnforced()
        {
            Assert.IsEmpty(Rules("Option Explicit", new SyncConfig()));
        }

        [Test]
        public void MissingAttribute_WhenNameIsMissingOrWrong()
        {
            var config = Enforcing("MissingAttribute");

            CollectionAssert.AreEqual(new[] { "MissingAttribute" }, Rules("Option Explicit", config));
            CollectionAssert.AreEqual(new[] { "MissingAttribute" }, Rules("Attribute VB_Name = \"Bar\"\r\nOption Explicit", config));
            Assert.IsEmpty(Rules("Attribute VB_Name = \"Foo\"\r\nOption Explicit", config));
        }

        [Test]
        public void ModuleWithoutFolder_UnlessThePathProvidesIt()
        {
            var config = Enforcing("ModuleWithoutFolder");
            var text = "Attribute VB_Name = \"Foo\"\r\nOption Explicit";

            CollectionAssert.AreEqual(new[] { "ModuleWithoutFolder" }, Rules(text, config));
            Assert.IsEmpty(Rules(text, config, folderFollowsPath: true));
            Assert.IsEmpty(Rules("Attribute VB_Name = \"Foo\"\r\n'@Folder(\"A\")\r\nOption Explicit", config));
        }

        [Test]
        public void MissingModuleAnnotation_WhenAttributeHasNoAnnotation()
        {
            var config = Enforcing("MissingModuleAnnotation");

            CollectionAssert.AreEqual(
                new[] { "MissingModuleAnnotation", "MissingModuleAnnotation" },
                Rules("Attribute VB_Name = \"Foo\"\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\nOption Explicit", config));
            Assert.IsEmpty(Rules("Attribute VB_Name = \"Foo\"\r\nAttribute VB_PredeclaredId = True\r\n'@PredeclaredId\r\nOption Explicit", config));
        }

        [Test]
        public void AttributeValueOutOfSync_WhenAnnotationHasNoAttribute()
        {
            var config = Enforcing("AttributeValueOutOfSync");

            CollectionAssert.AreEqual(new[] { "AttributeValueOutOfSync" }, Rules("Attribute VB_Name = \"Foo\"\r\n'@PredeclaredId\r\nOption Explicit", config));
            CollectionAssert.AreEqual(new[] { "AttributeValueOutOfSync" }, Rules("Attribute VB_Name = \"Foo\"\r\n'@ModuleAttribute VB_Ext_Key, \"Key\", \"Value\"\r\nOption Explicit", config));
            Assert.IsEmpty(Rules("Attribute VB_Name = \"Foo\"\r\nAttribute VB_Ext_Key = \"Key\", \"Value\"\r\n'@ModuleAttribute VB_Ext_Key, \"Key\", \"Value\"\r\nOption Explicit", config));
        }

        [Test]
        public void Violations_AreFixableOnlyInFixMode()
        {
            var config = Enforcing("MissingAttribute", "ModuleWithoutFolder");
            config.OnViolation = ViolationMode.Fix;

            var violations = SyncHygiene.Check("Foo", "Option Explicit", SyncSide.Module, config).ToList();

            Assert.IsTrue(violations.Single(v => v.Rule == "MissingAttribute").Fixable);
            Assert.IsFalse(violations.Single(v => v.Rule == "ModuleWithoutFolder").Fixable);
            Assert.IsFalse(violations.Any(v => v.Blocking));
        }

        [Test]
        public void Violations_BlockOnlyInBlockMode()
        {
            var config = Enforcing("MissingAttribute");
            Assert.IsTrue(SyncHygiene.Check("Foo", "Option Explicit", SyncSide.Module, config).Single().Blocking);

            config.OnViolation = ViolationMode.Warn;
            Assert.IsFalse(SyncHygiene.Check("Foo", "Option Explicit", SyncSide.Module, config).Single().Blocking);
        }
    }

    [TestFixture]
    [Category("GitSync")]
    public class GitSyncPlannerTests
    {
        private static readonly IIndenterSettings Settings = new IndenterSettings(true);
        private static readonly ISimpleIndenter Indenter = new Indenter(null, () => Settings);

        private const string Original = "Attribute VB_Name = \"Foo\"\r\nOption Explicit\r\nPublic Sub A()\r\nEnd Sub\r\n";
        private const string Changed = "Attribute VB_Name = \"Foo\"\r\nOption Explicit\r\nPublic Sub B()\r\nEnd Sub\r\n";
        private const string ChangedElsewhere = "Attribute VB_Name = \"Foo\"\r\nOption Explicit\r\nPublic Sub C()\r\nEnd Sub\r\n";

        private static string Sig(string text) => SyncText.Signature(text, Settings, Indenter);

        private static ModuleInput Module(string text, string name = "Foo") => new ModuleInput { Name = name, Extension = ".cls", ExportedText = text };
        private static FileInput File(string text, string name = "Foo", string path = "Foo.cls") => new FileInput { Name = name, RelativePath = path, Text = text };

        private static ModuleBaseline Baseline(string moduleText, string fileText, string path = "Foo.cls") =>
            new ModuleBaseline { Module = Sig(moduleText), File = Sig(fileText), Path = path, Folder = string.Empty };

        private static SyncPlan Plan(ModuleInput[] modules, FileInput[] files, ModuleBaseline baseline = null, SyncConfig config = null, GitIgnoreSet ignored = null)
        {
            var state = new GitSyncState();
            if (baseline != null)
            {
                state.Modules["Foo"] = baseline;
            }

            return new GitSyncPlanner(Settings, Indenter, config).Plan(modules, files, state, ignored ?? GitIgnoreSet.Empty);
        }

        [Test]
        public void IdenticalContent_IsInSync_EvenWithOtherLineEndings()
        {
            var plan = Plan(new[] { Module(Original) }, new[] { File(Original.Replace("\r\n", "\n")) });

            Assert.IsEmpty(plan.Items);
            Assert.AreEqual(1, plan.InSyncCount);
            Assert.IsTrue(plan.InSyncBaselines.ContainsKey("Foo"));
        }

        [Test]
        public void ModuleWithoutFile_IsExported()
        {
            var item = Plan(new[] { Module(Original) }, new FileInput[0]).Items.Single();

            Assert.AreEqual(SyncStatus.ModuleOnly, item.Status);
            Assert.AreEqual(SyncAction.Export, item.Suggested);
            Assert.AreEqual("Foo.cls", item.RelativePath);
        }

        [Test]
        public void ModuleWithoutFile_GoesToItsFolder()
        {
            var text = "Attribute VB_Name = \"Foo\"\r\n'@Folder(\"Models.Orders\")\r\nOption Explicit\r\n";

            var item = Plan(new[] { Module(text) }, new FileInput[0]).Items.Single();

            Assert.AreEqual("Models\\Orders\\Foo.cls", item.RelativePath);
        }

        [Test]
        public void FileWithoutModule_IsImported()
        {
            var item = Plan(new ModuleInput[0], new[] { File(Original) }).Items.Single();

            Assert.AreEqual(SyncStatus.FileOnly, item.Status);
            Assert.AreEqual(SyncAction.Import, item.Suggested);
        }

        [Test]
        public void OnlyModuleChanged_SuggestsExport()
        {
            var item = Plan(new[] { Module(Changed) }, new[] { File(Original) }, Baseline(Original, Original)).Items.Single();

            Assert.AreEqual(SyncStatus.ModuleChanged, item.Status);
            Assert.AreEqual(SyncAction.Export, item.Suggested);
        }

        [Test]
        public void OnlyFileChanged_SuggestsImport()
        {
            var item = Plan(new[] { Module(Original) }, new[] { File(Changed) }, Baseline(Original, Original)).Items.Single();

            Assert.AreEqual(SyncStatus.FileChanged, item.Status);
            Assert.AreEqual(SyncAction.Import, item.Suggested);
        }

        [Test]
        public void BothChanged_NeedsADecision()
        {
            var item = Plan(new[] { Module(Changed) }, new[] { File(ChangedElsewhere) }, Baseline(Original, Original)).Items.Single();

            Assert.AreEqual(SyncStatus.BothChanged, item.Status);
            Assert.AreEqual(SyncAction.None, item.Suggested);
        }

        [Test]
        public void DifferentWithoutHistory_NeedsADecision()
        {
            var item = Plan(new[] { Module(Changed) }, new[] { File(Original) }).Items.Single();

            Assert.AreEqual(SyncStatus.Differs, item.Status);
            Assert.AreEqual(SyncAction.None, item.Suggested);
        }

        [Test]
        public void FormattingOnlyChange_IsNotAChange()
        {
            var reindented = Original.Replace("Public Sub A()\r\n", "Public Sub A()\r\n\r\n");
            var plan = Plan(new[] { Module(Original) }, new[] { File(Original.Replace("Option Explicit", "    Option Explicit")) }, Baseline(Original, reindented));

            Assert.IsEmpty(plan.Items);
        }

        [Test]
        public void ModuleMovedToAnotherFolder_MovesItsFile()
        {
            var moved = "Attribute VB_Name = \"Foo\"\r\n'@Folder(\"New\")\r\nOption Explicit\r\nPublic Sub A()\r\nEnd Sub\r\n";
            var fileInOldPlace = "Attribute VB_Name = \"Foo\"\r\n'@Folder(\"Old\")\r\nOption Explicit\r\nPublic Sub A()\r\nEnd Sub\r\n";
            var baseline = new ModuleBaseline { Module = Sig(fileInOldPlace), File = Sig(fileInOldPlace), Path = "Old/Foo.cls", Folder = "Old" };

            var item = Plan(new[] { Module(moved) }, new[] { File(fileInOldPlace, path: "Old\\Foo.cls") }, baseline).Items.Single();

            Assert.IsTrue(item.MovesFile);
            Assert.AreEqual(SyncAction.Export, item.Suggested);
        }

        [Test]
        public void IgnoredItems_AreLeftOut()
        {
            var ignored = new GitIgnoreSet(new[] { new System.Collections.Generic.KeyValuePair<string, GitIgnore>(string.Empty, new GitIgnore(new[] { "/Foo.cls" })) });

            var plan = Plan(new[] { Module(Changed) }, new[] { File(Original) }, ignored: ignored);

            Assert.IsEmpty(plan.Items);
            CollectionAssert.AreEqual(new[] { "Foo" }, plan.Ignored);
        }

        [Test]
        public void ViolationsBlockTheSideThatBreaksTheRule()
        {
            var config = new SyncConfig { OnViolation = ViolationMode.Block };
            config.Enforce.Add("ModuleWithoutFolder");
            var withFolder = "Attribute VB_Name = \"Foo\"\r\n'@Folder(\"A\")\r\nOption Explicit\r\n";

            var item = Plan(new[] { Module(Changed) }, new[] { File(withFolder) }, config: config).Items.Single();

            Assert.IsTrue(item.IsBlockedFor(SyncAction.Export), "the module has no @Folder");
            Assert.IsFalse(item.IsBlockedFor(SyncAction.Import), "the file has one");
        }

        [Test]
        public void ChosenBlockedAction_IsReportedAsBlocked()
        {
            var config = new SyncConfig { OnViolation = ViolationMode.Block };
            config.Enforce.Add("MissingAttribute");

            var item = Plan(new[] { Module("Option Explicit\r\n") }, new FileInput[0], config: config).Items.Single();
            item.Chosen = SyncAction.Export;

            Assert.IsTrue(item.IsBlocked);
        }
    }
}
