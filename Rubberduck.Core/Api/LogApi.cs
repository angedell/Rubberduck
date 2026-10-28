using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Rubberduck.Api
{
    /// <summary>The logging switch of Rubberduck, as far as the API needs it. Replaceable so tests need not touch the real settings.</summary>
    public interface ILogControl
    {
        /// <summary>The level in the settings: Trace, Debug, Info, Warn, Error, Fatal or Off.</summary>
        string ConfiguredLevel { get; }

        /// <summary>The level logging actually runs at right now. Differs from the configured one when the setting was not applied.</summary>
        string EffectiveLevel { get; }

        /// <summary>The file the log is written to.</summary>
        string LogFile { get; }

        /// <summary>
        /// Saves the level in the settings as the user's own choice (so it is not reset to Off when Rubberduck shuts down) and applies it at once.
        /// </summary>
        void SetLevel(string level);
    }

    /// <summary>Rubberduck's own log, readable without finding the file, and its level, changeable without opening the settings.</summary>
    public class LogApi
    {
        public const int DefaultTail = 200;
        public const int MaxTail = 2000;

        //How much of the end of the file is looked at. The log rolls over into archives, so this is the recent past.
        private const int MaxBytesRead = 4 * 1024 * 1024;

        /// <summary>Most to least verbose, as Rubberduck's settings name them. Off writes nothing.</summary>
        public static readonly IReadOnlyList<string> Levels = new[] { "Trace", "Debug", "Info", "Warn", "Error", "Fatal", "Off" };

        //A new entry starts with the time; "2026-10-27 06:28:22.9468;INFO-2.5.92.13;Rubberduck._Extension;message;".
        private static readonly Regex EntryStart = new Regex(@"^(?<time>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}(?:\.\d+)?);(?<level>[A-Za-z]+)(?:-[^;]*)?;(?<logger>[^;]*);(?<message>.*)$", RegexOptions.Compiled);

        private readonly ILogControl _control;
        private readonly IFileSystem _fileSystem;
        private readonly ApiWriteApproval _approval;

        public LogApi(ILogControl control, IFileSystem fileSystem, ApiWriteApproval approval)
        {
            _control = control;
            _fileSystem = fileSystem;
            _approval = approval;
        }

        /// <param name="tail">How many of the latest matching entries to return. Default 200, at most 2000.</param>
        /// <param name="minLevel">Leave out entries less severe than this.</param>
        /// <param name="contains">Only entries whose message or logger contains this text (any case).</param>
        public LogReport Get(int? tail, string minLevel, string contains, string logger)
        {
            if (tail < 0 || tail > MaxTail)
            {
                throw ApiException.BadRequest($"tail must be between 0 and {MaxTail}.");
            }

            var minimum = 0;
            if (!string.IsNullOrWhiteSpace(minLevel))
            {
                minimum = IndexOfLevel(minLevel, "minLevel");
                if (minimum == Levels.Count - 1)
                {
                    throw ApiException.BadRequest("minLevel cannot be Off; pick one of: " + string.Join(", ", Levels.Take(Levels.Count - 1)) + ".");
                }
            }

            var report = new LogReport
            {
                ConfiguredLevel = _control.ConfiguredLevel,
                EffectiveLevel = _control.EffectiveLevel,
                Levels = Levels.ToList(),
                Path = _control.LogFile
            };

            if (!string.Equals(report.ConfiguredLevel, report.EffectiveLevel, StringComparison.OrdinalIgnoreCase))
            {
                report.Note = $"The settings say {report.ConfiguredLevel}, but logging runs at {report.EffectiveLevel}: the setting has not taken effect. Set the level again, or restart Rubberduck.";
            }
            else if (string.Equals(report.EffectiveLevel, "Off", StringComparison.OrdinalIgnoreCase))
            {
                report.Note = "Logging is off, so nothing new is written. Set the level to Debug or Trace to record what Rubberduck does.";
            }

            if (!_fileSystem.File.Exists(report.Path))
            {
                report.Note = (report.Note == null ? string.Empty : report.Note + " ") + "There is no log file yet.";
                return report;
            }

            var entries = Parse(ReadEnd(report.Path, report));
            var matching = entries
                .Where(e => IndexOfLevelOrZero(e.Level) >= minimum)
                .Where(e => string.IsNullOrEmpty(contains) || Contains(e.Message, contains) || Contains(e.Logger, contains))
                .Where(e => string.IsNullOrEmpty(logger) || Contains(e.Logger, logger))
                .ToList();

            var take = tail ?? DefaultTail;
            report.Total = matching.Count;
            report.Truncated = matching.Count > take;
            report.Entries = matching.Skip(Math.Max(0, matching.Count - take)).ToList();
            return report;
        }

        /// <summary>Sets the level. It is a change to the user's settings, so it needs the same approval as writes.</summary>
        public LogReport SetLevel(string client, string level)
        {
            if (string.IsNullOrWhiteSpace(level))
            {
                throw ApiException.BadRequest("'level' is required: " + string.Join(", ", Levels) + ".");
            }

            var name = Levels[IndexOfLevel(level, "level")];
            _approval.Require(client);
            _control.SetLevel(name);

            var report = Get(0, null, null, null);
            if (name == "Trace" || name == "Debug")
            {
                report.Note = (report.Note == null ? string.Empty : report.Note + " ") + "Verbose logging slows the editor a little and the log grows quickly; set it back when you are done.";
            }

            return report;
        }

        private static int IndexOfLevel(string level, string field)
        {
            var at = Levels.ToList().FindIndex(l => l.Equals(level.Trim(), StringComparison.OrdinalIgnoreCase));
            return at >= 0 ? at : throw ApiException.BadRequest($"{field} must be one of: {string.Join(", ", Levels)}.");
        }

        //The file names levels in capitals ("WARN"); one it does not know counts as the most verbose, so it is never hidden by accident.
        private static int IndexOfLevelOrZero(string level)
        {
            var at = Levels.ToList().FindIndex(l => l.Equals(level, StringComparison.OrdinalIgnoreCase));
            return at < 0 ? 0 : at;
        }

        private static bool Contains(string text, string part) => text != null && text.IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0;

        //The logger keeps the file open for writing, so it is read in a way that allows that.
        private string ReadEnd(string path, LogReport report)
        {
            using (var stream = _fileSystem.File.Open(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite | System.IO.FileShare.Delete))
            {
                report.SizeBytes = stream.Length;
                var skipped = stream.Length > MaxBytesRead;
                if (skipped)
                {
                    stream.Seek(-MaxBytesRead, System.IO.SeekOrigin.End);
                }

                using (var reader = new System.IO.StreamReader(stream, Encoding.UTF8, true))
                {
                    var text = reader.ReadToEnd();
                    if (skipped)
                    {
                        //Started in the middle of a line.
                        var firstBreak = text.IndexOf('\n');
                        text = firstBreak < 0 ? string.Empty : text.Substring(firstBreak + 1);
                        report.StartsMidFile = true;
                    }

                    return text;
                }
            }
        }

        /// <summary>Splits the log into entries. A line that does not start with a time belongs to the entry before it (a stack trace, a message with line breaks).</summary>
        public static List<LogEntry> Parse(string text)
        {
            var entries = new List<LogEntry>();
            LogEntry current = null;
            foreach (var raw in text.Replace("\r\n", "\n").Split('\n'))
            {
                var match = EntryStart.Match(raw);
                if (match.Success)
                {
                    current = new LogEntry
                    {
                        Time = match.Groups["time"].Value,
                        Level = Capitalize(match.Groups["level"].Value),
                        Logger = match.Groups["logger"].Value,
                        Message = match.Groups["message"].Value
                    };
                    entries.Add(current);
                }
                else if (current != null && raw.Length > 0)
                {
                    current.Message += "\n" + raw;
                }
            }

            foreach (var entry in entries)
            {
                //Every message ends with the format's own separator.
                entry.Message = entry.Message.TrimEnd().TrimEnd(';');
            }

            return entries;
        }

        private static string Capitalize(string level) => level.Length == 0 ? level : char.ToUpperInvariant(level[0]) + level.Substring(1).ToLowerInvariant();
    }
}
