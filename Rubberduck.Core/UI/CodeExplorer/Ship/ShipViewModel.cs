using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Threading.Tasks;
using Rubberduck.Resources.Ship;
using Rubberduck.Ship;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Ship
{
    public enum ShipPage { Settings, Confirm, Working, Done }

    public enum ShipFolderState { NotSet, Missing, NotWritable, Ok }

    public enum ShipPreviewState { NotRun, Running, Ready, Stale, Failed }

    /// <summary>What the dialog needs to know about the disk. A seam so the view model is tested without touching it.</summary>
    public interface IShipFileProbe
    {
        /// <param name="probeWrite">False while the user is typing: only whether the folder is there, without writing a probe file into it.</param>
        ShipFolderState CheckFolder(string folder, bool probeWrite = true);
        bool FileExists(string path);
    }

    public sealed class ShipFileProbe : IShipFileProbe
    {
        private readonly IFileSystem _fileSystem;

        public ShipFileProbe(IFileSystem fileSystem = null)
        {
            _fileSystem = fileSystem ?? new FileSystem();
        }

        public ShipFolderState CheckFolder(string folder, bool probeWrite = true)
        {
            if (string.IsNullOrWhiteSpace(folder))
            {
                return ShipFolderState.NotSet;
            }

            try
            {
                if (!_fileSystem.Directory.Exists(folder))
                {
                    return ShipFolderState.Missing;
                }

                if (!probeWrite)
                {
                    return ShipFolderState.Ok;
                }

                var probe = _fileSystem.Path.Combine(folder, ".rubberduck-ship-" + Guid.NewGuid().ToString("N") + ".tmp");
                using (_fileSystem.FileStream.Create(probe, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write, System.IO.FileShare.None, 1, System.IO.FileOptions.DeleteOnClose))
                {
                }

                return ShipFolderState.Ok;
            }
            catch (Exception)
            {
                return ShipFolderState.NotWritable;
            }
        }

        public bool FileExists(string path) => !string.IsNullOrEmpty(path) && _fileSystem.File.Exists(path);
    }

    /// <summary>
    /// Everything the Ship dialog decides: the choices, their validation, which buttons work, the confirmation step and the run.
    /// No controls in here; the dialog reads it and calls it. <see cref="Changed"/> is raised after every change, possibly from the thread that ships.
    /// </summary>
    public sealed class ShipViewModel
    {
        private static readonly IPath Paths = new FileSystem().Path;

        private readonly IShipService _service;
        private readonly IVBProject _project;
        private readonly IShipFileProbe _probe;
        private readonly object _gate = new object();
        private readonly Dictionary<string, bool> _steps = ShipSteps.All.ToDictionary(step => step, step => !ShipSteps.OffByDefault(step));
        private readonly List<string> _progress = new List<string>();
        private List<string> _keep = new List<string>();
        private ShipSettings _loaded = new ShipSettings();
        private ShipOnExisting? _answer;
        private int _previewVersion;

        public ShipViewModel(IShipService service, IVBProject project, string projectName, IShipFileProbe probe = null)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
            _project = project;
            ProjectName = projectName ?? string.Empty;
            _probe = probe ?? new ShipFileProbe();
        }

        public event EventHandler Changed;

        public string ProjectName { get; }

        public ShipPage Page { get; private set; } = ShipPage.Settings;

        // --- choices

        public ShipFormat Format { get; private set; } = ShipFormat.Xlsm;

        /// <summary>The folder as written in the settings.</summary>
        public string OutputFolder { get; private set; } = string.Empty;

        /// <summary>What the user typed; empty means the project's name.</summary>
        public string FileName { get; private set; } = string.Empty;

        public ShipOnExisting OnExisting { get; private set; } = ShipOnExisting.Ask;

        public bool RememberExistingChoice { get; private set; }

        /// <summary>The user's switch: ship without asking. Applied when the user ships.</summary>
        public bool Unattended { get; private set; }

        public bool IsStepOn(string step) => _steps.TryGetValue(step, out var on) ? on : !ShipSteps.OffByDefault(step);

        public ShipTreeShakingMode TreeShakingMode { get; private set; } = ShipTreeShakingMode.Conservative;

        /// <summary>The names the user wants kept whatever tree shaking finds: "Member" or "Module.Member".</summary>
        public IReadOnlyList<string> TreeShakingKeep => _keep;

        /// <summary>The keep list as the dialog's box shows it, one name per line.</summary>
        public string TreeShakingKeepText => string.Join(Environment.NewLine, _keep);

        // --- derived

        /// <summary>The folder the settings resolve to on this machine, or null.</summary>
        public string ResolvedFolder { get; private set; }

        public ShipFolderState FolderState { get; private set; } = ShipFolderState.NotSet;

        public bool FileNameValid => FileName.IndexOfAny(Paths.GetInvalidFileNameChars()) < 0;

        public string Extension => Format == ShipFormat.Xlam ? ".xlam" : ".xlsm";

        public string EffectiveFileName => (string.IsNullOrWhiteSpace(FileName) ? ProjectName : FileName.Trim()) + Extension;

        /// <summary>Where the file will be written, if the folder and the name are usable.</summary>
        public string TargetPath
        {
            get
            {
                if (string.IsNullOrWhiteSpace(ResolvedFolder) || !FileNameValid)
                {
                    return null;
                }

                try
                {
                    return Paths.Combine(ResolvedFolder, EffectiveFileName);
                }
                catch (ArgumentException)
                {
                    return null;
                }
            }
        }

        public bool TargetExists => TargetPath != null && _probe.FileExists(TargetPath);

        /// <summary>The file is in the way and the settings say "ask": the confirmation needs the answer.</summary>
        public bool NeedsExistingAnswer => OnExisting == ShipOnExisting.Ask && TargetExists;

        // --- preview

        public ShipPreviewState PreviewState { get; private set; } = ShipPreviewState.NotRun;

        public ShipReport PreviewReport { get; private set; }

        public string PreviewError { get; private set; }

        /// <summary>Set when the preview was refused for a reason Ship cannot get past by itself (unsaved workbook, not Excel, Excel unavailable...).</summary>
        public string BlockedMessage { get; private set; }

        public string BlockedReason { get; private set; }

        // --- state of the run

        public string ResultError { get; private set; }

        public ShipReport Result { get; private set; }

        public IReadOnlyList<string> ProgressLines
        {
            get
            {
                lock (_gate)
                {
                    return _progress.ToArray();
                }
            }
        }

        /// <summary>The folder to open from the result, once something was written.</summary>
        public string ResultFolder
        {
            get
            {
                var path = Result?.OutputPath;
                return string.IsNullOrEmpty(path) ? null : Paths.GetDirectoryName(path);
            }
        }

        /// <summary>Whether the Ship button of the settings page works: valid folder and name, nothing refused, no preview running.</summary>
        public bool CanShip =>
            Page == ShipPage.Settings
            && FolderState == ShipFolderState.Ok
            && FileNameValid
            && BlockedMessage == null
            && PreviewState != ShipPreviewState.Running;

        public bool CanGoBack => Page == ShipPage.Confirm;

        /// <summary>On the confirmation page, the button that writes. When the file exists and nothing was decided beforehand, there are two (replace / keep both).</summary>
        public bool CanConfirm => Page == ShipPage.Confirm && !NeedsExistingAnswer;

        // --- loading and editing

        public void Load()
        {
            var settings = _service.ReadSettings(_project) ?? new ShipSettings();
            _loaded = settings;
            Format = settings.Format;
            OutputFolder = settings.OutputFolder ?? string.Empty;
            FileName = settings.FileName ?? string.Empty;
            OnExisting = settings.OnExisting;
            foreach (var step in ShipSteps.All)
            {
                _steps[step] = settings.IsOn(step);
            }

            TreeShakingMode = settings.TreeShaking.Mode;
            _keep = settings.TreeShaking.Keep.ToList();

            Unattended = _service.IsUnattended(_project);
            Revalidate();
        }

        public void SetFormat(ShipFormat format)
        {
            if (Format == format)
            {
                return;
            }

            Format = format;
            Edited();
        }

        public void SetOutputFolder(string folder)
        {
            OutputFolder = folder ?? string.Empty;
            Edited();
        }

        public void SetFileName(string name)
        {
            FileName = StripExtension(name ?? string.Empty);
            Edited();
        }

        public void SetStep(string step, bool on)
        {
            _steps[step] = on;
            MarkPreviewStale();
            RaiseChanged();
        }

        public void SetTreeShakingMode(ShipTreeShakingMode mode)
        {
            if (TreeShakingMode == mode)
            {
                return;
            }

            TreeShakingMode = mode;
            MarkPreviewStale();
            RaiseChanged();
        }

        /// <summary>Takes the box's text: one name per line, blank lines and repeats (any case) dropped.</summary>
        public void SetTreeShakingKeepText(string text)
        {
            var names = (text ?? string.Empty)
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(line => line.Trim())
                .Where(line => line.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (names.SequenceEqual(_keep))
            {
                return;
            }

            _keep = names;
            MarkPreviewStale();
            RaiseChanged();
        }

        public void SetOnExisting(ShipOnExisting value)
        {
            OnExisting = value;
            RaiseChanged();
        }

        public void SetRememberExistingChoice(bool value)
        {
            RememberExistingChoice = value;
            RaiseChanged();
        }

        public void SetUnattended(bool value)
        {
            Unattended = value;
            RaiseChanged();
        }

        // --- preview

        /// <summary>Runs the pipeline on the copy without writing, off the calling thread. A result that is overtaken by a newer request is dropped.</summary>
        public async Task RefreshPreviewAsync()
        {
            int version;
            ShipSettings settings;
            lock (_gate)
            {
                version = ++_previewVersion;
                PreviewState = ShipPreviewState.Running;
                PreviewError = null;
                BlockedMessage = null;
                BlockedReason = null;
                settings = BuildSettings(OnExisting);
            }

            RaiseChanged();

            ShipReport report = null;
            ShipRefusedException refused = null;
            Exception failure = null;
            try
            {
                report = await Task.Run(() => _service.Preview(_project, settings)).ConfigureAwait(false);
            }
            catch (ShipRefusedException exception)
            {
                refused = exception;
            }
            catch (Exception exception)
            {
                failure = exception;
            }

            lock (_gate)
            {
                if (version != _previewVersion)
                {
                    return;
                }

                PreviewReport = report;
                if (refused != null)
                {
                    PreviewState = ShipPreviewState.Failed;
                    PreviewError = refused.Message;
                    if (!IsFolderReason(refused.Reason))
                    {
                        BlockedReason = refused.Reason;
                        BlockedMessage = refused.Message;
                    }
                }
                else if (failure != null)
                {
                    PreviewState = ShipPreviewState.Failed;
                    PreviewError = failure.Message;
                }
                else
                {
                    PreviewState = ShipPreviewState.Ready;
                }
            }

            RaiseChanged();
        }

        // --- the flow

        /// <summary>Whether pressing Ship shows the confirmation first. Only the user's unattended switch (as the service knows it) skips it, and a pending existing-file question cannot be skipped.</summary>
        public bool NeedsConfirmation => !_service.IsUnattended(_project) || NeedsExistingAnswer;

        /// <summary>The Ship button of the settings page: goes to the confirmation, or straight to the run when the user is unattended.</summary>
        public Task ProceedAsync()
        {
            //Typing only checked that the folder is there; here a file is really made in it.
            Revalidate();
            if (!CanShip)
            {
                return Task.CompletedTask;
            }

            if (NeedsConfirmation)
            {
                Page = ShipPage.Confirm;
                RaiseChanged();
                return Task.CompletedTask;
            }

            return RunAsync();
        }

        public void Back()
        {
            if (Page != ShipPage.Confirm)
            {
                return;
            }

            Page = ShipPage.Settings;
            RaiseChanged();
        }

        /// <summary>The write button of the confirmation page. <paramref name="answer"/> is needed (replace or keep both) when the file exists and the settings say ask.</summary>
        public Task ConfirmAsync(ShipOnExisting? answer = null)
        {
            if (Page != ShipPage.Confirm)
            {
                return Task.CompletedTask;
            }

            if (NeedsExistingAnswer)
            {
                if (answer == null || answer == ShipOnExisting.Ask)
                {
                    return Task.CompletedTask;
                }

                _answer = answer;
            }

            return RunAsync();
        }

        private async Task RunAsync()
        {
            var existing = OnExisting != ShipOnExisting.Ask ? OnExisting : (_answer ?? ShipOnExisting.Ask);
            var settings = BuildSettings(RememberExistingChoice && existing != ShipOnExisting.Ask ? existing : _loaded.OnExisting);
            var progress = new LineProgress(this);

            lock (_gate)
            {
                _progress.Clear();
            }

            Page = ShipPage.Working;
            Result = null;
            ResultError = null;
            RaiseChanged();

            try
            {
                if (Unattended != _service.IsUnattended(_project))
                {
                    _service.SetUnattended(_project, Unattended);
                }

                _service.WriteSettings(_project, settings);
                Result = await Task.Run(() => _service.Ship(_project, settings, existing, progress)).ConfigureAwait(false);
            }
            catch (ShipRefusedException exception)
            {
                ResultError = exception.Message;
                BlockedReason = exception.Reason;
            }
            catch (Exception exception)
            {
                ResultError = exception.Message;
            }

            Page = ShipPage.Done;
            RaiseChanged();
        }

        /// <summary>The settings as the dialog has them now; <paramref name="onExisting"/> is what goes into the file.</summary>
        public ShipSettings BuildSettings(ShipOnExisting onExisting)
        {
            var settings = new ShipSettings
            {
                Format = Format,
                OutputFolder = string.IsNullOrWhiteSpace(OutputFolder) ? null : OutputFolder.Trim(),
                FileName = string.IsNullOrWhiteSpace(FileName) ? null : FileName.Trim(),
                OnExisting = onExisting
            };
            foreach (var step in ShipSteps.All)
            {
                settings.Steps[step] = IsStepOn(step);
            }

            settings.TreeShaking.Mode = TreeShakingMode;
            settings.TreeShaking.Keep.AddRange(_keep);

            return settings;
        }

        // --- text for the summary (the same lines are on the settings page and in the confirmation)

        public IReadOnlyList<string> PreviewLines()
        {
            switch (PreviewState)
            {
                case ShipPreviewState.NotRun:
                    return new[] { ShipUI.Ship_Preview_NotRun };
                case ShipPreviewState.Running:
                    return new[] { ShipUI.Ship_Preview_Running };
                case ShipPreviewState.Failed:
                    var lines = new List<string> { string.Format(ShipUI.Ship_Preview_Failed, PreviewError) };
                    if (BlockedReason == "unsaved")
                    {
                        lines.Add(ShipUI.Ship_Refused_Unsaved_Hint);
                    }

                    return lines;
                default:
                    var report = DescribeReport(PreviewReport).ToList();
                    if (PreviewState == ShipPreviewState.Stale)
                    {
                        report.Insert(0, ShipUI.Ship_Preview_Stale);
                    }

                    return report;
            }
        }

        public static IEnumerable<string> DescribeReport(ShipReport report)
        {
            if (report == null)
            {
                yield break;
            }

            var removed = report.Steps.Where(s => s.Removed > 0).ToList();
            if (removed.Count == 0)
            {
                yield return ShipUI.Ship_Preview_NothingRemoved;
            }

            foreach (var step in removed)
            {
                yield return string.Format(ShipUI.Ship_Preview_StepLine, StepLabel(step.Step), step.Removed);
                foreach (var detail in step.Details)
                {
                    yield return "    " + detail;
                }
            }

            if (report.DynamicCalls.Count > 0)
            {
                yield return ShipUI.Ship_Preview_DynamicCalls;
                foreach (var call in report.DynamicCalls)
                {
                    yield return "    " + call;
                }
            }

            foreach (var unparsed in report.Unparsed)
            {
                yield return string.Format(ShipUI.Ship_Preview_Unparsed, unparsed);
            }

            foreach (var warning in report.Warnings)
            {
                yield return string.Format(ShipUI.Ship_Preview_Warning, warning);
            }
        }

        public static string StepLabel(string step)
        {
            return ShipUI.ResourceManager.GetString("Ship_Step_" + step, ShipUI.Culture) ?? step;
        }

        public IReadOnlyList<string> ConfirmationLines()
        {
            var removed = ShipSteps.All.Where(IsStepOn).Select(StepLabel).ToList();
            var lines = new List<string>
            {
                string.Format(ShipUI.Ship_Confirm_File, TargetPath),
                string.Format(ShipUI.Ship_Confirm_Type, Extension),
                string.Format(ShipUI.Ship_Confirm_Removes, removed.Count == 0 ? ShipUI.Ship_Confirm_RemovesNothing : string.Join("; ", removed))
            };
            if (IsStepOn(ShipSteps.TreeShaking))
            {
                lines.Add(string.Format(ShipUI.Ship_Confirm_TreeShaking,
                    TreeShakingMode == ShipTreeShakingMode.EntryPointsOnly ? ShipUI.Ship_TreeShaking_Mode_EntryPointsOnly : ShipUI.Ship_TreeShaking_Mode_Conservative,
                    _keep.Count));
            }

            lines.Add(ShipUI.Ship_Confirm_Untouched);
            return lines;
        }

        public string FolderMessage
        {
            get
            {
                switch (FolderState)
                {
                    case ShipFolderState.NotSet: return ShipUI.Ship_FolderNotSet;
                    case ShipFolderState.Missing: return string.Format(ShipUI.Ship_FolderMissing, ResolvedFolder);
                    case ShipFolderState.NotWritable: return string.Format(ShipUI.Ship_FolderNotWritable, ResolvedFolder);
                    default: return null;
                }
            }
        }

        // --- internals

        private void Edited()
        {
            MarkPreviewStale();
            Revalidate(probeWrite: false);
        }

        private void MarkPreviewStale()
        {
            lock (_gate)
            {
                if (PreviewState == ShipPreviewState.Ready || PreviewState == ShipPreviewState.Failed)
                {
                    PreviewState = ShipPreviewState.Stale;
                }
            }
        }

        private void Revalidate(bool probeWrite = true)
        {
            try
            {
                ResolvedFolder = _service.ResolveOutputFolder(_project, BuildSettings(OnExisting));
            }
            catch (Exception)
            {
                ResolvedFolder = null;
            }

            FolderState = _probe.CheckFolder(ResolvedFolder, probeWrite);
            RaiseChanged();
        }

        private static string StripExtension(string name)
        {
            var trimmed = name.Trim();
            foreach (var extension in new[] { ".xlsm", ".xlam" })
            {
                if (trimmed.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
                {
                    return trimmed.Substring(0, trimmed.Length - extension.Length);
                }
            }

            return name;
        }

        private static bool IsFolderReason(string reason) => reason == "folderMissing" || reason == "folderNotWritable";

        private void RaiseChanged() => Changed?.Invoke(this, EventArgs.Empty);

        private sealed class LineProgress : IProgress<string>
        {
            private readonly ShipViewModel _owner;

            public LineProgress(ShipViewModel owner) { _owner = owner; }

            public void Report(string value)
            {
                lock (_owner._gate)
                {
                    _owner._progress.Add(value);
                }

                _owner.RaiseChanged();
            }
        }
    }
}
