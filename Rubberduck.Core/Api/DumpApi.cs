using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Abstractions;
using System.Linq;
using System.Threading;

namespace Rubberduck.Api
{
    /// <summary>Starts a program and waits for it. Replaceable so tests need not run one.</summary>
    public interface IProcessRunner
    {
        /// <returns>The exit code and what the program wrote, or null for the exit code if it did not finish in time (it is stopped then).</returns>
        (int? ExitCode, string Output) Run(string file, string arguments, TimeSpan timeout);
    }

    public class ProcessRunner : IProcessRunner
    {
        public (int? ExitCode, string Output) Run(string file, string arguments, TimeSpan timeout)
        {
            var info = new ProcessStartInfo(file, arguments) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
            using (var process = Process.Start(info))
            {
                var output = process.StandardOutput.ReadToEndAsync();
                var errors = process.StandardError.ReadToEndAsync();
                if (!process.WaitForExit((int)timeout.TotalMilliseconds))
                {
                    try { process.Kill(); } catch (InvalidOperationException) { }
                    return (null, output.Result + errors.Result);
                }

                return (process.ExitCode, output.Result + errors.Result);
            }
        }
    }

    /// <summary>A program left running in the background, ProcDump watching the editor for instance.</summary>
    public interface IBackgroundProcess
    {
        bool HasExited { get; }

        int? ExitCode { get; }

        /// <summary>What the program has written so far (the end of it).</summary>
        string Output { get; }

        void Kill();
    }

    public interface IProcessLauncher
    {
        IBackgroundProcess Start(string file, string arguments);
    }

    public class ProcessLauncher : IProcessLauncher
    {
        private const int MaxOutput = 4000;

        private class Running : IBackgroundProcess
        {
            private readonly Process _process;
            private readonly System.Text.StringBuilder _output = new System.Text.StringBuilder();

            public Running(Process process)
            {
                _process = process;
                DataReceivedEventHandler collect = (sender, e) =>
                {
                    if (e.Data == null)
                    {
                        return;
                    }

                    lock (_output)
                    {
                        _output.AppendLine(e.Data);
                        if (_output.Length > MaxOutput)
                        {
                            _output.Remove(0, _output.Length - MaxOutput);
                        }
                    }
                };
                process.OutputDataReceived += collect;
                process.ErrorDataReceived += collect;
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
            }

            public bool HasExited => _process.HasExited;

            public int? ExitCode => _process.HasExited ? _process.ExitCode : (int?)null;

            public string Output
            {
                get
                {
                    lock (_output)
                    {
                        return _output.ToString();
                    }
                }
            }

            public void Kill()
            {
                try
                {
                    if (!_process.HasExited)
                    {
                        _process.Kill();
                    }
                }
                catch (InvalidOperationException)
                {
                }
            }
        }

        public IBackgroundProcess Start(string file, string arguments)
        {
            var info = new ProcessStartInfo(file, arguments) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
            return new Running(Process.Start(info));
        }
    }

    public class DumpReport
    {
        /// <summary>Where ProcDump was found, or null.</summary>
        public string Tool { get; set; }

        /// <summary>Where dumps are written.</summary>
        public string Folder { get; set; }

        /// <summary>The dump just taken.</summary>
        public string File { get; set; }

        public long? SizeBytes { get; set; }

        /// <summary>Dumps already in the folder, newest first.</summary>
        public List<string> Existing { get; set; } = new List<string>();

        public string Note { get; set; }

        /// <summary>Set when the folder is one a cloud service syncs: a dump holds everything in memory and would be uploaded.</summary>
        public string Warning { get; set; }

        /// <summary>What the monitor is doing, when one was started.</summary>
        public DumpMonitorReport Monitor { get; set; }

        /// <summary>What is wrong, when the dump was analyzed on capture.</summary>
        public DumpAnalysis Analysis { get; set; }
    }

    public class DumpDeleteReport
    {
        public List<string> Deleted { get; set; } = new List<string>();

        public long FreedBytes { get; set; }

        /// <summary>Files that could not be deleted, with the reason (ProcDump may still be writing one).</summary>
        public List<string> Failed { get; set; } = new List<string>();

        public List<string> Existing { get; set; } = new List<string>();
    }

    public class DumpMonitorReport
    {
        public bool Running { get; set; }

        public int? ProcessId { get; set; }

        /// <summary>What the monitor waits for, as it was started.</summary>
        public string Triggers { get; set; }

        public DateTime? StartedUtc { get; set; }

        /// <summary>Dumps written since the monitor started, newest first.</summary>
        public List<string> NewDumps { get; set; } = new List<string>();

        /// <summary>The newest of them, read, when asked for.</summary>
        public DumpAnalysis Analysis { get; set; }

        /// <summary>The end of what ProcDump has written.</summary>
        public string Output { get; set; }

        public string Note { get; set; }
    }

    /// <summary>
    /// A memory dump of the editor's process, taken with Sysinternals ProcDump, for when the editor hangs: the dump says what every thread was doing,
    /// which neither the log nor the API can. The dump holds everything the process had in memory, the open documents included, so taking one needs the same
    /// approval as a write. It is taken from the API's own thread, so it works while the editor's UI thread is stuck.
    /// </summary>
    public class DumpApi
    {
        public const string ToolNotFound = "procdump_not_found";

        private static readonly TimeSpan DefaultWait = TimeSpan.FromSeconds(120);
        private static readonly TimeSpan MaxWait = TimeSpan.FromSeconds(300);

        private readonly ApiWriteApproval _approval;
        private readonly IFileSystem _fileSystem;
        private readonly IProcessRunner _runner;
        private readonly Func<int> _processId;
        private readonly Func<string, string> _environment;
        private readonly IDumpAnalyzer _analyzer;
        private readonly IProcessLauncher _launcher;
        private readonly Func<DateTime> _now;

        private static readonly TimeSpan AnalysisLimit = TimeSpan.FromMinutes(3);
        private static readonly string[] CloudFolders = { "dropbox", "onedrive", "google drive", "googledrive", "icloud", "box sync", "sync.com", "pcloud" };

        private readonly object _monitorLock = new object();
        private IBackgroundProcess _monitor;
        private string _monitorTriggers;
        private DateTime? _monitorStarted;
        private HashSet<string> _monitorBaseline = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public DumpApi(ApiWriteApproval approval, IFileSystem fileSystem, IProcessRunner runner = null, Func<int> processId = null, Func<string, string> environment = null,
            IDumpAnalyzer analyzer = null, IProcessLauncher launcher = null, Func<DateTime> now = null)
        {
            _approval = approval;
            _fileSystem = fileSystem;
            _runner = runner ?? new ProcessRunner();
            _processId = processId ?? (() => Process.GetCurrentProcess().Id);
            _environment = environment ?? Environment.GetEnvironmentVariable;
            _analyzer = analyzer ?? new ClrMdDumpAnalyzer();
            _launcher = launcher ?? new ProcessLauncher();
            _now = now ?? (() => DateTime.UtcNow);
        }

        /// <summary>RUBBERDUCK_DUMP_FOLDER if set (a shared folder, say, so the dump can be read from another machine), otherwise Rubberduck's Dumps folder under LOCALAPPDATA.</summary>
        private string Folder
        {
            get
            {
                var chosen = _environment("RUBBERDUCK_DUMP_FOLDER");
                return string.IsNullOrWhiteSpace(chosen)
                    ? _fileSystem.Path.Combine(_environment("LOCALAPPDATA") ?? string.Empty, "Rubberduck", "Dumps")
                    : chosen.Trim();
            }
        }

        /// <summary>ProcDump, looked for in RUBBERDUCK_PROCDUMP, in Rubberduck's own Tools folder and on the PATH.</summary>
        public string FindTool()
        {
            var candidates = new List<string> { _environment("RUBBERDUCK_PROCDUMP") };
            candidates.Add(_fileSystem.Path.Combine(_environment("LOCALAPPDATA") ?? string.Empty, "Rubberduck", "Tools", "procdump.exe"));
            foreach (var folder in (_environment("PATH") ?? string.Empty).Split(';').Where(f => f.Trim().Length > 0))
            {
                candidates.Add(_fileSystem.Path.Combine(folder.Trim(), "procdump.exe"));
            }

            return candidates.FirstOrDefault(c => !string.IsNullOrWhiteSpace(c) && _fileSystem.File.Exists(c));
        }

        /// <summary>Whether a dump can be taken, and what is already in the folder. Takes nothing.</summary>
        public DumpReport Status()
        {
            var report = new DumpReport { Tool = FindTool(), Folder = Folder, Existing = Existing(), Warning = SyncWarning(Folder) };
            if (report.Tool == null)
            {
                report.Note = NotFoundMessage();
            }

            var monitor = MonitorReport(false);
            report.Monitor = monitor.Running || monitor.NewDumps.Any() ? monitor : null;
            return report;
        }

        /// <summary>A dump holds everything the process had in memory; in a folder a cloud service syncs it leaves the machine.</summary>
        private string SyncWarning(string folder)
        {
            var lowered = (folder ?? string.Empty).ToLowerInvariant();
            var service = CloudFolders.FirstOrDefault(lowered.Contains);
            return service == null
                ? null
                : $"The dump folder looks like one that {service} syncs. A dump holds everything in memory, the open documents included, and would be uploaded. "
                  + "Pause the sync first, or point the RUBBERDUCK_DUMP_FOLDER environment variable at a folder outside it and restart the editor.";
        }

        /// <param name="kind">"full" (default: all memory, which is what reading managed stacks needs) or "mini" (small, threads and stacks only).</param>
        /// <param name="waitSeconds">How long to wait for ProcDump; default 120, at most 300.</param>
        /// <param name="analyze">Read the dump as soon as it is written, so the answer says what is wrong.</param>
        public DumpReport Capture(string client, string kind, int? waitSeconds, bool analyze = false)
        {
            var mode = string.IsNullOrWhiteSpace(kind) ? "full" : kind.Trim().ToLowerInvariant();
            if (mode != "full" && mode != "mini")
            {
                throw ApiException.BadRequest("kind must be 'full' or 'mini'.");
            }

            if (waitSeconds < 1 || waitSeconds > MaxWait.TotalSeconds)
            {
                throw ApiException.BadRequest($"waitSeconds must be between 1 and {MaxWait.TotalSeconds}.");
            }

            var tool = FindTool();
            if (tool == null)
            {
                throw new ApiException(404, ToolNotFound, NotFoundMessage());
            }

            _approval.Require(client);

            var folder = Folder;
            _fileSystem.Directory.CreateDirectory(folder);
            var before = new HashSet<string>(Existing(), StringComparer.OrdinalIgnoreCase);
            var wait = waitSeconds.HasValue ? TimeSpan.FromSeconds(waitSeconds.Value) : DefaultWait;
            var result = _runner.Run(tool, $"-accepteula {(mode == "full" ? "-ma" : "-mm")} {_processId()} \"{folder}\"", wait);

            var made = Existing().FirstOrDefault(f => !before.Contains(f));
            if (made == null)
            {
                throw ApiException.Unavailable(result.ExitCode == null
                    ? "ProcDump did not finish in time and was stopped; no dump was written. Try 'mini', or a longer waitSeconds."
                    : $"ProcDump wrote no dump (exit code {result.ExitCode}): {Tail(result.Output)}");
            }

            var report = new DumpReport
            {
                Tool = tool,
                Folder = folder,
                File = made,
                SizeBytes = _fileSystem.FileInfo.FromFileName(made).Length,
                Existing = Existing(),
                Warning = SyncWarning(folder),
                Note = result.ExitCode == null ? "ProcDump was still running when the wait ended and was stopped; the file may be incomplete." : null
            };

            if (analyze)
            {
                report.Analysis = Read(made);
            }

            return report;
        }

        /// <summary>
        /// Reads a dump and says what is wrong. Taking nothing from the editor and showing only stacks and exception messages, it needs no approval,
        /// unless the dump is to be deleted afterwards.
        /// </summary>
        /// <param name="name">A file name as GET /dump lists it, or "latest".</param>
        public DumpAnalysis Analyze(string client, string name, bool deleteAfter = false)
        {
            var file = Resolve(name);
            if (deleteAfter)
            {
                _approval.Require(client);
            }

            var analysis = Read(file);
            if (deleteAfter)
            {
                var report = Remove(new[] { file });
                analysis.Deleted = report.Deleted.Any();
                analysis.Note = report.Failed.Any()
                    ? (analysis.Note + " The dump could not be deleted: " + string.Join("; ", report.Failed)).Trim()
                    : analysis.Note;
            }

            return analysis;
        }

        private DumpAnalysis Read(string file)
        {
            using (var limit = new CancellationTokenSource(AnalysisLimit))
            {
                var analysis = _analyzer.Analyze(file, limit.Token);
                analysis.File = file;
                analysis.SizeBytes = _fileSystem.FileInfo.FromFileName(file).Length;
                return analysis;
            }
        }

        /// <summary>Deletes one dump. A dump is the editor's whole memory, so it is deleted when it has been read; this is how.</summary>
        /// <param name="name">A file name as GET /dump lists it, or "latest".</param>
        public DumpDeleteReport Delete(string client, string name)
        {
            var file = Resolve(name);
            _approval.Require(client);
            return Remove(new[] { file });
        }

        /// <summary>Deletes the dumps in the folder, the newest <paramref name="keep"/> of them excepted.</summary>
        public DumpDeleteReport DeleteAll(string client, int? keep)
        {
            if (keep < 0)
            {
                throw ApiException.BadRequest("keep cannot be negative.");
            }

            var doomed = Existing().Skip(keep ?? 0).ToList();
            if (doomed.Any())
            {
                _approval.Require(client);
            }

            return Remove(doomed);
        }

        private DumpDeleteReport Remove(IEnumerable<string> files)
        {
            var report = new DumpDeleteReport();
            foreach (var file in files)
            {
                try
                {
                    var size = _fileSystem.FileInfo.FromFileName(file).Length;
                    _fileSystem.File.Delete(file);
                    report.Deleted.Add(file);
                    report.FreedBytes += size;
                }
                catch (Exception exception) when (exception is System.IO.IOException || exception is UnauthorizedAccessException)
                {
                    report.Failed.Add(_fileSystem.Path.GetFileName(file) + ": " + exception.Message);
                }
            }

            report.Existing = Existing();
            return report;
        }

        /// <summary>Only files of the dump folder can be named, and by their name alone: no path gets the caller anywhere else.</summary>
        private string Resolve(string name)
        {
            var existing = Existing();
            var wanted = (name ?? string.Empty).Trim();
            if (wanted.Length == 0 || wanted.Equals("latest", StringComparison.OrdinalIgnoreCase))
            {
                return existing.FirstOrDefault() ?? throw ApiException.NotFound("There is no dump in " + Folder + ".");
            }

            if (wanted.IndexOfAny(new[] { '\\', '/', ':' }) >= 0 || !wanted.EndsWith(".dmp", StringComparison.OrdinalIgnoreCase))
            {
                throw ApiException.BadRequest("Name a dump by its file name alone, as GET /dump lists them (EXCEL.EXE_261106_052301.dmp), or 'latest'.");
            }

            return existing.FirstOrDefault(f => _fileSystem.Path.GetFileName(f).Equals(wanted, StringComparison.OrdinalIgnoreCase))
                   ?? throw ApiException.NotFound($"There is no dump named '{wanted}' in {Folder}.");
        }

        /// <summary>
        /// Has ProcDump watch the editor and write a dump by itself when it hangs (a window that stops answering), throws an unhandled exception or keeps a processor busy,
        /// so the moment is caught even when nobody is there to ask for it. It needs the approval a dump needs.
        /// </summary>
        /// <param name="hang">Dump when a window of the process does not answer for 5 seconds. On unless said otherwise.</param>
        /// <param name="exceptions">Dump on an unhandled exception.</param>
        /// <param name="cpuPercent">Dump when the process keeps this much of the processor (of one core) for 10 seconds.</param>
        /// <param name="maxDumps">How many dumps to write before ProcDump stops; default 2, at most 10.</param>
        public DumpMonitorReport StartMonitor(string client, string kind, bool? hang, bool? exceptions, int? cpuPercent, int? maxDumps)
        {
            var mode = string.IsNullOrWhiteSpace(kind) ? "full" : kind.Trim().ToLowerInvariant();
            if (mode != "full" && mode != "mini")
            {
                throw ApiException.BadRequest("kind must be 'full' or 'mini'.");
            }

            if (cpuPercent < 1 || cpuPercent > 100)
            {
                throw ApiException.BadRequest("cpuPercent must be between 1 and 100.");
            }

            if (maxDumps < 1 || maxDumps > 10)
            {
                throw ApiException.BadRequest("maxDumps must be between 1 and 10.");
            }

            var triggers = new List<string>();
            var arguments = new List<string> { "-accepteula", mode == "full" ? "-ma" : "-mm" };
            if (hang ?? (exceptions != true && cpuPercent == null))
            {
                arguments.Add("-h");
                triggers.Add("a hung window");
            }

            if (exceptions == true)
            {
                arguments.Add("-e");
                triggers.Add("an unhandled exception");
            }

            if (cpuPercent.HasValue)
            {
                arguments.Add($"-c {cpuPercent} -s 10");
                triggers.Add($"{cpuPercent}% processor for 10 seconds");
            }

            if (!triggers.Any())
            {
                throw ApiException.BadRequest("Nothing to watch for: set at least one of hang, exceptions and cpuPercent.");
            }

            var tool = FindTool();
            if (tool == null)
            {
                throw new ApiException(404, ToolNotFound, NotFoundMessage());
            }

            lock (_monitorLock)
            {
                if (_monitor != null && !_monitor.HasExited)
                {
                    throw ApiException.Conflict("A dump monitor is already running; stop it first (DELETE /dump/monitor).");
                }

                _approval.Require(client);

                var folder = Folder;
                _fileSystem.Directory.CreateDirectory(folder);
                arguments.Add($"-n {maxDumps ?? 2}");
                arguments.Add(_processId().ToString());
                arguments.Add($"\"{folder}\"");

                _monitorBaseline = new HashSet<string>(Existing(), StringComparer.OrdinalIgnoreCase);
                _monitorStarted = _now();
                _monitorTriggers = string.Join(", ", triggers);
                _monitor = _launcher.Start(tool, string.Join(" ", arguments));
            }

            var report = MonitorReport(false);
            var warning = SyncWarning(Folder);
            report.Note = warning;
            return report;
        }

        /// <summary>Whether the monitor runs and what it has caught.</summary>
        /// <param name="analyze">Read the newest dump it wrote, so the answer says what is wrong.</param>
        public DumpMonitorReport MonitorStatus(bool analyze) => MonitorReport(analyze);

        public DumpMonitorReport StopMonitor()
        {
            lock (_monitorLock)
            {
                _monitor?.Kill();
            }

            return MonitorReport(false);
        }

        private DumpMonitorReport MonitorReport(bool analyze)
        {
            var report = MonitorSnapshot();
            if (analyze && report.NewDumps.Any())
            {
                report.Analysis = Read(report.NewDumps.First());
            }

            return report;
        }

        private DumpMonitorReport MonitorSnapshot()
        {
            lock (_monitorLock)
            {
                if (_monitor == null)
                {
                    return new DumpMonitorReport { Note = "No monitor has been started. POST /dump/monitor starts one." };
                }

                var report = new DumpMonitorReport
                {
                    Running = !_monitor.HasExited,
                    Triggers = _monitorTriggers,
                    StartedUtc = _monitorStarted,
                    ProcessId = _processId(),
                    NewDumps = Existing().Where(f => !_monitorBaseline.Contains(f)).ToList(),
                    Output = Tail(_monitor.Output)
                };

                if (report.NewDumps.Any())
                {
                    report.Note = $"The monitor has caught {report.NewDumps.Count} dump{(report.NewDumps.Count == 1 ? string.Empty : "s")}.";
                }
                else if (!report.Running)
                {
                    report.Note = $"The monitor ended (exit code {_monitor.ExitCode}) without a dump.";
                }

                return report;
            }
        }

        private List<string> Existing()
        {
            var folder = Folder;
            if (!_fileSystem.Directory.Exists(folder))
            {
                return new List<string>();
            }

            return _fileSystem.Directory.GetFiles(folder, "*.dmp")
                .OrderByDescending(f => _fileSystem.File.GetLastWriteTimeUtc(f))
                .ToList();
        }

        private static string Tail(string text)
        {
            var trimmed = (text ?? string.Empty).Trim();
            return trimmed.Length <= 400 ? trimmed : trimmed.Substring(trimmed.Length - 400);
        }

        private string NotFoundMessage()
            => "ProcDump (Sysinternals, procdump.exe) was not found. Put it in " + _fileSystem.Path.Combine(_environment("LOCALAPPDATA") ?? string.Empty, "Rubberduck", "Tools")
               + ", add its folder to the PATH, or point the RUBBERDUCK_PROCDUMP environment variable at it. Download: https://learn.microsoft.com/sysinternals/downloads/procdump";
    }
}
