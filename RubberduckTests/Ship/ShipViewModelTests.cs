using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using NUnit.Framework;
using Rubberduck.Ship;
using Rubberduck.UI.CodeExplorer.Ship;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Ship
{
    [TestFixture]
    [Category("Ship")]
    public class ShipViewModelTests
    {
        internal class FakeShipService : IShipService
        {
            public ShipSettings Settings = new ShipSettings { OutputFolder = "out" };
            public bool Unattended;
            public bool CanShipResult = true;
            public string Resolved = @"C:\work\out";
            public Func<ShipSettings, ShipReport> OnPreview = s => new ShipReport { DryRun = true };
            public Func<ShipSettings, ShipOnExisting, IProgress<string>, ShipReport> OnShip =
                (s, e, p) => new ShipReport { OutputPath = @"C:\work\out\Book.xlsm" };

            public List<ShipSettings> Written = new List<ShipSettings>();
            public List<bool> UnattendedSet = new List<bool>();
            public List<ShipOnExisting> ShipCalls = new List<ShipOnExisting>();
            public int PreviewCalls;

            public bool CanShip(IVBProject project) => CanShipResult;
            public ShipSettings ReadSettings(IVBProject project) => Settings;
            public void WriteSettings(IVBProject project, ShipSettings settings) => Written.Add(settings);
            public string ResolveOutputFolder(IVBProject project, ShipSettings settings) => string.IsNullOrEmpty(settings.OutputFolder) ? null : Resolved;
            public ShipReport Preview(IVBProject project, ShipSettings settings) { PreviewCalls++; return OnPreview(settings); }

            public ShipReport Ship(IVBProject project, ShipSettings settings, ShipOnExisting onExisting, IProgress<string> progress)
            {
                ShipCalls.Add(onExisting);
                return OnShip(settings, onExisting, progress);
            }

            public bool IsUnattended(IVBProject project) => Unattended;
            public void SetUnattended(IVBProject project, bool value) { Unattended = value; UnattendedSet.Add(value); }
        }

        internal class FakeProbe : IShipFileProbe
        {
            public ShipFolderState Folder = ShipFolderState.Ok;
            public HashSet<string> Files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            public bool LastProbeWrite;
            public ShipFolderState CheckFolder(string folder, bool probeWrite = true) { LastProbeWrite = probeWrite; return Check(folder); }
            private ShipFolderState Check(string folder) => string.IsNullOrWhiteSpace(folder) ? ShipFolderState.NotSet : Folder;
            public bool FileExists(string path) => Files.Contains(path);
        }

        [Test]
        public void TypingOnlyChecksThatTheFolderIsThere_ShippingMakesTheProbeFile()
        {
            var probe = new FakeProbe();
            var model = Make(new FakeShipService { Settings = new ShipSettings { OutputFolder = "dist" }, Resolved = "C:\\dist" }, probe);

            model.SetOutputFolder("dist2");
            Assert.IsFalse(probe.LastProbeWrite);

            model.ProceedAsync().Wait();
            Assert.IsTrue(probe.LastProbeWrite);
        }

        private static ShipViewModel Make(FakeShipService service, FakeProbe probe = null)
        {
            var model = new ShipViewModel(service, null, "Book", probe ?? new FakeProbe());
            model.Load();
            return model;
        }

        [Test]
        public void ProjectWithNoSavedChoice_PreselectsXlsmAndTheStepsThatAreOnByDefault()
        {
            var model = Make(new FakeShipService());

            Assert.AreEqual(ShipFormat.Xlsm, model.Format);
            Assert.AreEqual(".xlsm", model.Extension);
            foreach (var step in ShipSteps.All)
            {
                Assert.AreEqual(!ShipSteps.OffByDefault(step), model.IsStepOn(step), step);
            }

            Assert.IsFalse(model.IsStepOn(ShipSteps.TreeShaking));
            Assert.AreEqual(ShipTreeShakingMode.Conservative, model.TreeShakingMode);
            Assert.AreEqual(0, model.TreeShakingKeep.Count);
        }

        [Test]
        public void Load_ShowsTheSavedTreeShakingModeAndKeepList()
        {
            var settings = new ShipSettings();
            settings.Steps[ShipSteps.TreeShaking] = true;
            settings.TreeShaking.Mode = ShipTreeShakingMode.EntryPointsOnly;
            settings.TreeShaking.Keep.AddRange(new[] { "Module1.Go", "Helper" });
            var model = Make(new FakeShipService { Settings = settings });

            Assert.IsTrue(model.IsStepOn(ShipSteps.TreeShaking));
            Assert.AreEqual(ShipTreeShakingMode.EntryPointsOnly, model.TreeShakingMode);
            CollectionAssert.AreEqual(new[] { "Module1.Go", "Helper" }, model.TreeShakingKeep);
        }

        [Test]
        public void SettingTheKeepListFromText_TakesOneNamePerLine_SkippingBlanksAndDuplicates()
        {
            var model = Make(new FakeShipService());

            model.SetTreeShakingKeepText("Module1.Go\r\n\n  Helper  \r\nmodule1.go\r\n");

            CollectionAssert.AreEqual(new[] { "Module1.Go", "Helper" }, model.TreeShakingKeep);
            var built = model.BuildSettings(ShipOnExisting.Ask);
            CollectionAssert.AreEqual(new[] { "Module1.Go", "Helper" }, built.TreeShaking.Keep);
        }

        [Test]
        public void BuildSettings_CarriesTheTreeShakingModeAndStepState()
        {
            var model = Make(new FakeShipService());
            model.SetStep(ShipSteps.TreeShaking, true);
            model.SetTreeShakingMode(ShipTreeShakingMode.EntryPointsOnly);

            var built = model.BuildSettings(ShipOnExisting.Ask);

            Assert.IsTrue(built.IsOn(ShipSteps.TreeShaking));
            Assert.AreEqual(ShipTreeShakingMode.EntryPointsOnly, built.TreeShaking.Mode);
        }

        [Test]
        public async Task ChangingTheModeOrTheKeepList_MakesThePreviewStale()
        {
            var model = Make(new FakeShipService());
            await model.RefreshPreviewAsync();
            model.SetTreeShakingMode(ShipTreeShakingMode.EntryPointsOnly);
            Assert.AreEqual(ShipPreviewState.Stale, model.PreviewState);

            await model.RefreshPreviewAsync();
            model.SetTreeShakingKeepText("Foo");
            Assert.AreEqual(ShipPreviewState.Stale, model.PreviewState);
        }

        [Test]
        public async Task Preview_ListsRemovalsByName_ThenTheDynamicCallsWithAnExplanation()
        {
            var report = new ShipReport { DryRun = true };
            var shaking = new ShipStepReport { Step = ShipSteps.TreeShaking, Removed = 2 };
            shaking.Details.Add("Module1.OldHelper");
            shaking.Details.Add("Module2.Unused");
            report.Steps.Add(shaking);
            report.DynamicCalls.Add("Module1.Go line 12: Application.Run with a computed name");
            var model = Make(new FakeShipService { OnPreview = s => report });
            await model.RefreshPreviewAsync();

            var lines = model.PreviewLines().ToList();
            var text = string.Join("\n", lines);

            StringAssert.Contains("Module1.OldHelper", text);
            StringAssert.Contains("Module2.Unused", text);
            StringAssert.Contains("Application.Run with a computed name", text);
            StringAssert.Contains("names to keep", text);
            Assert.Less(text.IndexOf("Module2.Unused"), text.IndexOf("Application.Run"));
        }

        [Test]
        public async Task Preview_WhenTreeShakingWasSkipped_ShowsTheReasonFromTheWarning()
        {
            var report = new ShipReport { DryRun = true };
            report.Warnings.Add("Tree shaking skipped: Module3 does not parse");
            var model = Make(new FakeShipService { OnPreview = s => report });
            await model.RefreshPreviewAsync();

            StringAssert.Contains("Tree shaking skipped: Module3 does not parse", string.Join("\n", model.PreviewLines()));
        }

        [Test]
        public void Confirmation_SaysTreeShakingIsOnOnlyWhenItIs()
        {
            var model = Make(new FakeShipService());
            StringAssert.DoesNotContain("tree shaking", string.Join("\n", model.ConfirmationLines()));

            model.SetStep(ShipSteps.TreeShaking, true);

            StringAssert.Contains("tree shaking", string.Join("\n", model.ConfirmationLines()));
        }

        [Test]
        public void Load_ShowsRememberedFolderAsWrittenAndTheResolvedOne()
        {
            var service = new FakeShipService { Settings = new ShipSettings { OutputFolder = @"..\dist", Format = ShipFormat.Xlam }, Resolved = @"C:\work\dist" };
            var model = Make(service);

            Assert.AreEqual(@"..\dist", model.OutputFolder);
            Assert.AreEqual(@"C:\work\dist", model.ResolvedFolder);
            Assert.AreEqual(ShipFormat.Xlam, model.Format);
            Assert.AreEqual(@"C:\work\dist\Book.xlam", model.TargetPath);
        }

        [Test]
        public void StepsSwitchedOffInTheSettings_AreOffInTheDialog()
        {
            var settings = new ShipSettings { OutputFolder = "out" };
            settings.Steps[ShipSteps.Comments] = false;
            var model = Make(new FakeShipService { Settings = settings });

            Assert.IsFalse(model.IsStepOn(ShipSteps.Comments));
            Assert.IsTrue(model.IsStepOn(ShipSteps.Notes));
        }

        [TestCase(ShipFolderState.Ok, true)]
        [TestCase(ShipFolderState.Missing, false)]
        [TestCase(ShipFolderState.NotWritable, false)]
        public void ShipIsEnabledOnlyForAUsableFolder(ShipFolderState state, bool enabled)
        {
            var model = Make(new FakeShipService(), new FakeProbe { Folder = state });

            Assert.AreEqual(state, model.FolderState);
            Assert.AreEqual(enabled, model.CanShip);
            Assert.AreEqual(enabled, model.FolderMessage == null);
        }

        [Test]
        public void NoFolderSet_DisablesShipAndAsksForOne()
        {
            var model = Make(new FakeShipService { Settings = new ShipSettings() });

            Assert.AreEqual(ShipFolderState.NotSet, model.FolderState);
            Assert.IsFalse(model.CanShip);
            Assert.IsNotNull(model.FolderMessage);
        }

        [Test]
        public void ChangingTheFolder_RevalidatesAndResolvesAgain()
        {
            var probe = new FakeProbe { Folder = ShipFolderState.Missing };
            var model = Make(new FakeShipService(), probe);
            Assert.IsFalse(model.CanShip);

            probe.Folder = ShipFolderState.Ok;
            model.SetOutputFolder("other");

            Assert.AreEqual("other", model.OutputFolder);
            Assert.IsTrue(model.CanShip);
        }

        [Test]
        public void FileNameWithBadCharacters_DisablesShip()
        {
            var model = Make(new FakeShipService());
            model.SetFileName("a<b");

            Assert.IsFalse(model.FileNameValid);
            Assert.IsFalse(model.CanShip);
            Assert.IsNull(model.TargetPath);
        }

        [Test]
        public void FileNameTypedWithExtension_KeepsTheNameOnly()
        {
            var model = Make(new FakeShipService());
            model.SetFileName("Tool.XLAM");

            Assert.AreEqual("Tool", model.FileName);
            Assert.AreEqual("Tool.xlsm", model.EffectiveFileName);
        }

        [Test]
        public void EmptyFileName_UsesTheProjectName()
        {
            var model = Make(new FakeShipService());

            Assert.AreEqual("Book.xlsm", model.EffectiveFileName);
        }

        [Test]
        public async Task Preview_ShowsStepsUnparsedAndWarningsWithoutBlocking()
        {
            var report = new ShipReport { DryRun = true };
            var notes = new ShipStepReport { Step = ShipSteps.Notes, Removed = 2 };
            notes.Details.Add("Notes1");
            report.Steps.Add(notes);
            report.Steps.Add(new ShipStepReport { Step = ShipSteps.Comments, Removed = 0 });
            report.Unparsed.Add("Broken: line 3");
            report.Warnings.Add("Pending resource changes");
            var service = new FakeShipService { OnPreview = s => report };
            var model = Make(service);

            await model.RefreshPreviewAsync();

            Assert.AreEqual(ShipPreviewState.Ready, model.PreviewState);
            var text = string.Join("\n", model.PreviewLines());
            StringAssert.Contains("Notes", text);
            StringAssert.Contains("Notes1", text);
            StringAssert.Contains("Broken: line 3", text);
            StringAssert.Contains("Pending resource changes", text);
            Assert.IsTrue(model.CanShip);
        }

        [Test]
        public async Task Preview_UsesTheStepsChosenInTheDialog()
        {
            ShipSettings seen = null;
            var service = new FakeShipService { OnPreview = s => { seen = s; return new ShipReport(); } };
            var model = Make(service);
            model.SetStep(ShipSteps.Comments, false);

            await model.RefreshPreviewAsync();

            Assert.IsFalse(seen.IsOn(ShipSteps.Comments));
            Assert.IsTrue(seen.IsOn(ShipSteps.Notes));
        }

        [Test]
        public async Task ChangingAStepAfterThePreview_MakesItStale()
        {
            var model = Make(new FakeShipService());
            await model.RefreshPreviewAsync();

            model.SetStep(ShipSteps.Notes, false);

            Assert.AreEqual(ShipPreviewState.Stale, model.PreviewState);
        }

        [Test]
        public async Task PreviewRefusedBecauseUnsaved_BlocksShipAndKeepsTheMessage()
        {
            var service = new FakeShipService { OnPreview = s => throw new ShipRefusedException("unsaved", "The workbook has unsaved changes.") };
            var model = Make(service);

            await model.RefreshPreviewAsync();

            Assert.AreEqual(ShipPreviewState.Failed, model.PreviewState);
            Assert.AreEqual("unsaved", model.BlockedReason);
            Assert.IsFalse(model.CanShip);
            var text = string.Join("\n", model.PreviewLines());
            StringAssert.Contains("unsaved changes", text);
            StringAssert.Contains("Rubberduck does not save", text);
        }

        [Test]
        public async Task PreviewRefusedForTheFolder_DoesNotBlockByItself()
        {
            var service = new FakeShipService { OnPreview = s => throw new ShipRefusedException("folderMissing", "No such folder.") };
            var model = Make(service);

            await model.RefreshPreviewAsync();

            Assert.IsNull(model.BlockedMessage);
            Assert.IsTrue(model.CanShip);
        }

        [Test]
        public async Task PreviewThatThrowsOtherwise_IsShownAndDoesNotBlock()
        {
            var service = new FakeShipService { OnPreview = s => throw new InvalidOperationException("boom") };
            var model = Make(service);

            await model.RefreshPreviewAsync();

            Assert.AreEqual(ShipPreviewState.Failed, model.PreviewState);
            StringAssert.Contains("boom", string.Join("\n", model.PreviewLines()));
            Assert.IsTrue(model.CanShip);
        }

        [Test]
        public async Task Proceed_GoesToConfirmationAndWritesNothing()
        {
            var service = new FakeShipService();
            var model = Make(service);

            await model.ProceedAsync();

            Assert.AreEqual(ShipPage.Confirm, model.Page);
            Assert.AreEqual(0, service.ShipCalls.Count);
            Assert.AreEqual(0, service.Written.Count);
            Assert.IsTrue(model.CanGoBack);
        }

        [Test]
        public async Task Back_ReturnsToSettings()
        {
            var model = Make(new FakeShipService());
            await model.ProceedAsync();

            model.Back();

            Assert.AreEqual(ShipPage.Settings, model.Page);
        }

        [Test]
        public async Task Proceed_WhenNotShippable_DoesNothing()
        {
            var model = Make(new FakeShipService(), new FakeProbe { Folder = ShipFolderState.Missing });

            await model.ProceedAsync();

            Assert.AreEqual(ShipPage.Settings, model.Page);
        }

        [Test]
        public async Task Confirm_WritesSettingsThenShipsAndReportsTheResult()
        {
            var service = new FakeShipService();
            var model = Make(service);
            model.SetFormat(ShipFormat.Xlam);
            model.SetStep(ShipSteps.Config, false);
            await model.ProceedAsync();

            await model.ConfirmAsync();

            Assert.AreEqual(ShipPage.Done, model.Page);
            Assert.AreEqual(1, service.Written.Count);
            Assert.AreEqual(ShipFormat.Xlam, service.Written[0].Format);
            Assert.IsFalse(service.Written[0].IsOn(ShipSteps.Config));
            Assert.AreEqual(1, service.ShipCalls.Count);
            Assert.AreEqual(@"C:\work\out\Book.xlsm", model.Result.OutputPath);
            Assert.AreEqual(@"C:\work\out", model.ResultFolder);
            Assert.IsNull(model.ResultError);
        }

        [Test]
        public async Task Unattended_SkipsTheConfirmation()
        {
            var service = new FakeShipService { Unattended = true };
            var model = Make(service);

            await model.ProceedAsync();

            Assert.AreEqual(ShipPage.Done, model.Page);
            Assert.AreEqual(1, service.ShipCalls.Count);
        }

        [Test]
        public async Task Unattended_StillAsksWhenTheFileExistsAndTheChoiceIsAsk()
        {
            var service = new FakeShipService { Unattended = true };
            var probe = new FakeProbe();
            probe.Files.Add(@"C:\work\out\Book.xlsm");
            var model = Make(service, probe);

            await model.ProceedAsync();

            Assert.AreEqual(ShipPage.Confirm, model.Page);
            Assert.AreEqual(0, service.ShipCalls.Count);
        }

        [Test]
        public async Task TickingShipWithoutAsking_AppliesOnShipAndStillConfirmsThisTime()
        {
            var service = new FakeShipService();
            var model = Make(service);
            model.SetUnattended(true);
            Assert.AreEqual(0, service.UnattendedSet.Count, "the switch is not applied by ticking it");

            await model.ProceedAsync();
            Assert.AreEqual(ShipPage.Confirm, model.Page);

            await model.ConfirmAsync();

            CollectionAssert.AreEqual(new[] { true }, service.UnattendedSet);
        }

        [Test]
        public async Task UntickingShipWithoutAsking_SwitchesItOff()
        {
            var service = new FakeShipService { Unattended = true };
            var probe = new FakeProbe();
            probe.Files.Add(@"C:\work\out\Book.xlsm");
            var model = Make(service, probe);
            model.SetUnattended(false);

            await model.ProceedAsync();
            await model.ConfirmAsync(ShipOnExisting.Replace);

            CollectionAssert.AreEqual(new[] { false }, service.UnattendedSet);
        }

        [Test]
        public async Task ExistingFileWithAsk_RequiresAnAnswer()
        {
            var service = new FakeShipService();
            var probe = new FakeProbe();
            probe.Files.Add(@"C:\work\out\Book.xlsm");
            var model = Make(service, probe);
            await model.ProceedAsync();

            Assert.IsTrue(model.NeedsExistingAnswer);
            Assert.IsFalse(model.CanConfirm);
            await model.ConfirmAsync();
            Assert.AreEqual(ShipPage.Confirm, model.Page, "no answer, nothing happens");
            Assert.AreEqual(0, service.ShipCalls.Count);

            await model.ConfirmAsync(ShipOnExisting.KeepBoth);

            CollectionAssert.AreEqual(new[] { ShipOnExisting.KeepBoth }, service.ShipCalls);
        }

        [Test]
        public async Task ExistingFileWithAStandingChoice_DoesNotAsk()
        {
            var service = new FakeShipService { Settings = new ShipSettings { OutputFolder = "out", OnExisting = ShipOnExisting.Replace } };
            var probe = new FakeProbe();
            probe.Files.Add(@"C:\work\out\Book.xlsm");
            var model = Make(service, probe);
            await model.ProceedAsync();

            Assert.IsFalse(model.NeedsExistingAnswer);
            await model.ConfirmAsync();

            CollectionAssert.AreEqual(new[] { ShipOnExisting.Replace }, service.ShipCalls);
        }

        [Test]
        public async Task RememberChoice_WritesItToTheSettings()
        {
            var service = new FakeShipService();
            var model = Make(service);
            model.SetOnExisting(ShipOnExisting.KeepBoth);
            model.SetRememberExistingChoice(true);
            await model.ProceedAsync();
            await model.ConfirmAsync();

            Assert.AreEqual(ShipOnExisting.KeepBoth, service.Written.Single().OnExisting);
        }

        [Test]
        public async Task ChoiceWithoutRemember_LeavesTheSettingsAsTheyWere()
        {
            var service = new FakeShipService { Settings = new ShipSettings { OutputFolder = "out", OnExisting = ShipOnExisting.Ask } };
            var model = Make(service);
            model.SetOnExisting(ShipOnExisting.Replace);
            await model.ProceedAsync();
            await model.ConfirmAsync();

            Assert.AreEqual(ShipOnExisting.Ask, service.Written.Single().OnExisting);
            Assert.AreEqual(ShipOnExisting.Replace, service.ShipCalls.Single());
        }

        [Test]
        public async Task AnswerAtTheConfirmation_IsRememberedWhenTicked()
        {
            var service = new FakeShipService();
            var probe = new FakeProbe();
            probe.Files.Add(@"C:\work\out\Book.xlsm");
            var model = Make(service, probe);
            model.SetRememberExistingChoice(true);
            await model.ProceedAsync();

            await model.ConfirmAsync(ShipOnExisting.Replace);

            Assert.AreEqual(ShipOnExisting.Replace, service.Written.Single().OnExisting);
        }

        [Test]
        public async Task Progress_IsCollectedWhileShipping()
        {
            var service = new FakeShipService
            {
                OnShip = (s, e, progress) =>
                {
                    progress.Report("Copying");
                    progress.Report("Saving");
                    return new ShipReport { OutputPath = @"C:\work\out\Book.xlsm" };
                }
            };
            var model = Make(service);
            await model.ProceedAsync();

            await model.ConfirmAsync();

            CollectionAssert.AreEqual(new[] { "Copying", "Saving" }, model.ProgressLines);
        }

        [Test]
        public async Task Shipping_RunsOffTheCallingThread()
        {
            var callerThread = Environment.CurrentManagedThreadId;
            var shipThread = -1;
            var service = new FakeShipService { OnShip = (s, e, p) => { shipThread = Environment.CurrentManagedThreadId; return new ShipReport(); } };
            var model = Make(service);
            await model.ProceedAsync();

            await model.ConfirmAsync();

            Assert.AreNotEqual(callerThread, shipThread);
        }

        [Test]
        public async Task RefusedShip_ShowsItsMessage()
        {
            var service = new FakeShipService { OnShip = (s, e, p) => throw new ShipRefusedException("unsaved", "Save the workbook first.") };
            var model = Make(service);
            await model.ProceedAsync();

            await model.ConfirmAsync();

            Assert.AreEqual(ShipPage.Done, model.Page);
            Assert.AreEqual("Save the workbook first.", model.ResultError);
            Assert.AreEqual("unsaved", model.BlockedReason);
            Assert.IsNull(model.ResultFolder);
        }

        [Test]
        public async Task UnexpectedFailure_ShowsTheMessageInsteadOfThrowing()
        {
            var service = new FakeShipService { OnShip = (s, e, p) => throw new InvalidOperationException("disk gone") };
            var model = Make(service);
            await model.ProceedAsync();

            await model.ConfirmAsync();

            Assert.AreEqual(ShipPage.Done, model.Page);
            Assert.AreEqual("disk gone", model.ResultError);
        }

        [Test]
        public async Task Confirmation_ListsFileTypeAndWhatIsRemoved()
        {
            var model = Make(new FakeShipService());
            model.SetStep(ShipSteps.Comments, false);
            await model.ProceedAsync();

            var text = string.Join("\n", model.ConfirmationLines());

            StringAssert.Contains(@"C:\work\out\Book.xlsm", text);
            StringAssert.Contains(".xlsm", text);
            StringAssert.DoesNotContain("Comments, including", text);
            StringAssert.Contains("Notes", text);
        }
    }
}
