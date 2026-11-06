using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Git
{
    public class GitResult
    {
        public GitResult(int exitCode, byte[] output, string error)
        {
            ExitCode = exitCode;
            OutputBytes = output;
            Error = error;
        }

        public int ExitCode { get; }
        public bool Ok => ExitCode == 0;

        /// <summary>What git wrote to standard output, as it is. Files can be in the ANSI code page, so this is not text until the caller knows its encoding.</summary>
        public byte[] OutputBytes { get; }

        public string Error { get; }

        /// <summary>The output read as UTF-8, which is what git's own messages, names and logs are.</summary>
        public string Output => Encoding.UTF8.GetString(OutputBytes);
    }

    public class GitException : Exception
    {
        public GitException(string message, int exitCode) : base(message)
        {
            ExitCode = exitCode;
        }

        public int ExitCode { get; }
    }

    /// <summary>
    /// Runs git as a child process: an argument list rather than a command line a shell would read, no prompts, no window, a time limit.
    /// A secret it was told about never reaches an error message or the log.
    /// </summary>
    public class GitRunner
    {
        private readonly string _git;
        private readonly TimeSpan _timeout;
        private readonly IReadOnlyDictionary<string, string> _environment;
        private readonly IReadOnlyList<string> _secrets;

        public GitRunner(string gitPath, TimeSpan? timeout = null, IReadOnlyDictionary<string, string> environment = null, IEnumerable<string> secrets = null)
        {
            _git = gitPath;
            _timeout = timeout ?? TimeSpan.FromSeconds(60);
            _environment = environment ?? new Dictionary<string, string>();
            _secrets = (secrets ?? Enumerable.Empty<string>()).Where(secret => !string.IsNullOrEmpty(secret)).ToList();
        }

        /// <summary>Runs git and returns what happened, whatever the exit code. Throws only if git cannot be started or does not finish in time.</summary>
        /// <exception cref="GitException">Git did not start, or took longer than the time limit and was stopped.</exception>
        public GitResult Run(IEnumerable<string> arguments, string workingDirectory = null)
        {
            var args = arguments.ToList();
            var info = new ProcessStartInfo(_git, CommandLine(args))
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                RedirectStandardInput = true,
                StandardErrorEncoding = Encoding.UTF8,
                WorkingDirectory = workingDirectory ?? Environment.CurrentDirectory
            };

            //Nobody is there to answer a question: git has to fail instead of waiting for one.
            info.EnvironmentVariables["GIT_TERMINAL_PROMPT"] = "0";
            info.EnvironmentVariables["GCM_INTERACTIVE"] = "never";
            info.EnvironmentVariables["GIT_OPTIONAL_LOCKS"] = "0";
            info.EnvironmentVariables["GIT_EDITOR"] = "true";
            info.EnvironmentVariables["LC_ALL"] = "C.UTF-8";
            foreach (var pair in _environment)
            {
                info.EnvironmentVariables[pair.Key] = pair.Value;
            }

            Process process;
            try
            {
                process = Process.Start(info);
            }
            catch (System.ComponentModel.Win32Exception exception)
            {
                throw new GitException($"git could not be started ({Mask(_git)}): {exception.Message}", -1);
            }

            using (process)
            {
                process.StandardInput.Close();
                var output = Task.Run(() =>
                {
                    using (var memory = new System.IO.MemoryStream())
                    {
                        process.StandardOutput.BaseStream.CopyTo(memory);
                        return memory.ToArray();
                    }
                });
                var error = process.StandardError.ReadToEndAsync();

                if (!process.WaitForExit((int)_timeout.TotalMilliseconds))
                {
                    try { process.Kill(); } catch (InvalidOperationException) { }
                    throw new GitException($"git {Mask(Describe(args))} did not finish in {_timeout.TotalSeconds:0} seconds and was stopped.", -1);
                }

                process.WaitForExit();
                return new GitResult(process.ExitCode, output.Result, Mask(error.Result.Trim()));
            }
        }

        /// <summary>Like <see cref="Run"/>, but a failing git is an exception that carries what git said.</summary>
        public GitResult RunOrThrow(IEnumerable<string> arguments, string workingDirectory = null)
        {
            var args = arguments.ToList();
            var result = Run(args, workingDirectory);
            if (!result.Ok)
            {
                throw new GitException($"git {Mask(Describe(args))} failed ({result.ExitCode}): {result.Error}", result.ExitCode);
            }

            return result;
        }

        private string Mask(string text)
            => _secrets.Aggregate(text ?? string.Empty, (current, secret) => current.Replace(secret, "***"));

        //Only for messages: the first words say what was asked without dragging a whole option list along.
        private static string Describe(IReadOnlyList<string> args)
        {
            var words = new List<string>();
            for (var i = 0; i < args.Count && words.Count < 2; i++)
            {
                if (args[i] == "-c")
                {
                    i++;
                }
                else if (!args[i].StartsWith("--git-dir") && !args[i].StartsWith("--work-tree"))
                {
                    words.Add(args[i]);
                }
            }

            return string.Join(" ", words);
        }

        /// <summary>
        /// Builds a Windows command line that the C runtime splits back into exactly these arguments:
        /// quotes around anything with a space or quote, and backslashes doubled where they come before a quote.
        /// </summary>
        public static string CommandLine(IEnumerable<string> arguments) => string.Join(" ", arguments.Select(Quote));

        public static string Quote(string argument)
        {
            if (argument.Length > 0 && argument.IndexOfAny(new[] { ' ', '\t', '\n', '\v', '"' }) < 0)
            {
                return argument;
            }

            var quoted = new StringBuilder("\"");
            var backslashes = 0;
            foreach (var c in argument)
            {
                if (c == '\\')
                {
                    backslashes++;
                    continue;
                }

                if (c == '"')
                {
                    quoted.Append('\\', backslashes * 2 + 1).Append('"');
                }
                else
                {
                    quoted.Append('\\', backslashes).Append(c);
                }

                backslashes = 0;
            }

            return quoted.Append('\\', backslashes * 2).Append('"').ToString();
        }
    }
}
