using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Abstractions;
using System.Linq;

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

        public DumpApi(ApiWriteApproval approval, IFileSystem fileSystem, IProcessRunner runner = null, Func<int> processId = null, Func<string, string> environment = null)
        {
            _approval = approval;
            _fileSystem = fileSystem;
            _runner = runner ?? new ProcessRunner();
            _processId = processId ?? (() => Process.GetCurrentProcess().Id);
            _environment = environment ?? Environment.GetEnvironmentVariable;
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
            var report = new DumpReport { Tool = FindTool(), Folder = Folder, Existing = Existing() };
            if (report.Tool == null)
            {
                report.Note = NotFoundMessage();
            }

            return report;
        }

        /// <param name="kind">"full" (default: all memory, which is what reading managed stacks needs) or "mini" (small, threads and stacks only).</param>
        /// <param name="waitSeconds">How long to wait for ProcDump; default 120, at most 300.</param>
        public DumpReport Capture(string client, string kind, int? waitSeconds)
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

            return new DumpReport
            {
                Tool = tool,
                Folder = folder,
                File = made,
                SizeBytes = _fileSystem.FileInfo.FromFileName(made).Length,
                Existing = Existing(),
                Note = result.ExitCode == null ? "ProcDump was still running when the wait ended and was stopped; the file may be incomplete." : null
            };
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
