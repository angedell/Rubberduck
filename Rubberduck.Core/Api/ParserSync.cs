using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using NLog;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Api
{
    /// <summary>
    /// What everything that reads the parser's results needs first: the project the caller means, and the parser caught up with the editor.
    /// Edits made through the API reach the parser only when somebody asks for a parse, so this asks, and waits.
    /// </summary>
    public class ParserSync
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static readonly TimeSpan ParseTimeout = TimeSpan.FromSeconds(90);
        private static readonly TimeSpan StartTimeout = TimeSpan.FromSeconds(2);

        private readonly RubberduckParserState _state;
        private readonly IParseManager _parseManager;
        private readonly IProjectsProvider _projects;
        private readonly IUiDispatcher _dispatcher;
        private readonly IVBE _vbe;

        public ParserSync(RubberduckParserState state, IParseManager parseManager, IProjectsProvider projects, IUiDispatcher dispatcher, IVBE vbe)
        {
            _state = state;
            _parseManager = parseManager;
            _projects = projects;
            _dispatcher = dispatcher;
            _vbe = vbe;
            //Every call of CatchUp ends in a parse run, which ends in Ready even when nothing was edited, so counting runs would say "changed" every time.
            //What counts is a module being scheduled for a parse (new, marked as modified) or actually parsed: a module whose text changed goes straight to Parsing, never through Pending.
            _state.ModuleStateChanged += (sender, args) =>
            {
                if (args.State == ParserState.Pending || args.State == ParserState.Parsing)
                {
                    Interlocked.Increment(ref _scheduled);
                }
            };
        }

        private int _scheduled;

        /// <summary>
        /// Changes when the code the parser holds may have changed: a module was scheduled for a parse, or the number of declarations differs (a module was removed).
        /// Results worked out from the parser state are good for as long as this stays the same.
        /// </summary>
        public int Generation => unchecked((Volatile.Read(ref _scheduled) * 397) ^ _state.AllUserDeclarations.Count());

        public RubberduckParserState State => _state;

        /// <summary>False while a macro runs or is stopped: the parser cannot make progress then, however long one waits.</summary>
        public bool IsEditorInDesignMode()
        {
            var task = _dispatcher.StartTask(() => _vbe.IsInDesignMode, CancellationToken.None);
            if (!task.Wait(TimeSpan.FromSeconds(10)))
            {
                throw ApiException.Unavailable("The editor did not respond in time.");
            }

            return task.Result;
        }

        public const string ModeDesign = "design";
        public const string ModeRunningOrBreak = "running_or_break";
        public const string ModeUnresponsive = "unresponsive";

        /// <summary>
        /// Where the editor is, for a caller to see before it tries anything: "design", "running_or_break" (a macro runs or is stopped; the editor cannot tell the two apart),
        /// or "unresponsive" (it did not answer in a few seconds: a dialog, a hung call). Never throws.
        /// </summary>
        public string EditorMode()
        {
            var task = _dispatcher.StartTask(() => _vbe.IsInDesignMode, CancellationToken.None);
            try
            {
                if (!task.Wait(TimeSpan.FromSeconds(5)))
                {
                    return ModeUnresponsive;
                }

                return task.Result ? ModeDesign : ModeRunningOrBreak;
            }
            catch (AggregateException)
            {
                return ModeUnresponsive;
            }
        }

        //The id of Run > Reset on the editor's menu.
        private const int ResetCommandId = 228;

        /// <summary>
        /// Presses Run > Reset if code is stopped (break mode), as after a compile error whose dialog was closed with End: the editor stays on the faulty line,
        /// and the next edit would raise a "this action will reset your project" prompt that nobody can answer for the caller. True if it is in design mode afterwards.
        /// </summary>
        public bool ResetIfStopped()
        {
            var task = _dispatcher.StartTask(() =>
            {
                if (_vbe.IsInDesignMode)
                {
                    return true;
                }

                using (var bars = _vbe.CommandBars)
                using (var reset = bars.FindControl(ResetCommandId))
                {
                    reset?.Execute();
                }

                return _vbe.IsInDesignMode;
            }, CancellationToken.None);

            return task.Wait(TimeSpan.FromSeconds(10)) && task.Result;
        }

        public (string Id, string Name) ResolveProject(string key)
        {
            var task = _dispatcher.StartTask(() => _projects.Projects().Select(p => (p.ProjectId, p.Project.Name, Path: ProjectIds.PathOf(p.Project))).ToList(), CancellationToken.None);
            if (!task.Wait(TimeSpan.FromSeconds(30)))
            {
                throw ApiException.Unavailable("The editor did not respond in time.");
            }

            var all = task.Result;
            var matches = ProjectIds.Match(key, all, p => (p.ProjectId, p.Name, p.Path));
            if (matches.Count == 0)
            {
                throw ApiException.NotFound($"No open project '{key}'.");
            }

            if (matches.Count > 1)
            {
                throw ApiException.Conflict($"More than one open project is named '{key}'; use its id.");
            }

            return (matches[0].ProjectId, matches[0].Name);
        }

        public bool ModuleExists(string projectId, string module)
            => _state.AllUserDeclarations.Any(d => string.Equals(d.QualifiedName.QualifiedModuleName.ProjectId, projectId, StringComparison.OrdinalIgnoreCase)
                                                   && string.Equals(d.QualifiedName.QualifiedModuleName.ComponentName, module, StringComparison.OrdinalIgnoreCase));

        //What the editor held when the parser last finished well, and how many modules had been scheduled for a parse by then.
        private ParseFingerprint _lastParsed;
        private int _lastParsedScheduled;

        private ParseFingerprint Capture()
        {
            var task = _dispatcher.StartTask(() =>
            {
                var modules = _projects.Components()
                    .Select(c => new KeyValuePair<string, int>($"{c.QualifiedModuleName.ProjectName}.{c.QualifiedModuleName.ComponentName}", QualifiedModuleName.GetContentHash(c.Component)))
                    .ToList();

                var references = new List<string>();
                foreach (var (projectId, project) in _projects.Projects())
                {
                    using (var projectReferences = project.References)
                    {
                        foreach (var reference in projectReferences)
                        {
                            using (reference)
                            {
                                references.Add($"{projectId}|{reference.Name}|{reference.Guid}|{reference.Major}.{reference.Minor}|{reference.FullPath}|{reference.IsBroken}");
                            }
                        }
                    }
                }

                return new ParseFingerprint(modules, references);
            }, CancellationToken.None);

            if (!task.Wait(TimeSpan.FromSeconds(30)))
            {
                throw ApiException.Unavailable("The editor did not respond in time.");
            }

            return task.Result;
        }

        /// <summary>
        /// Why a parse is needed, or null if it is not: the parser is Ready, nothing was scheduled for a parse since it last finished, and the editor holds the same modules and references.
        /// A run with nothing to parse still takes seconds on a large project (it settles libraries and resolves references again), so it is worth not asking for one.
        /// </summary>
        private string WhyParseIsNeeded(ParseFingerprint now)
        {
            if (_lastParsed == null)
            {
                return "no parse of the current state is known to have finished well";
            }

            if (_state.Status != ParserState.Ready)
            {
                return $"the parser is {_state.Status}";
            }

            if (_lastParsedScheduled != Volatile.Read(ref _scheduled))
            {
                return "modules were scheduled for a parse since the last one finished";
            }

            var differences = now.DifferencesFrom(_lastParsed);
            return differences.Count == 0 ? null : string.Join("; ", differences);
        }

        /// <summary>
        /// Asks for a parse, waits until the parser is done, and says how it ended. Unless <paramref name="force"/> is set, it asks for none
        /// when the parser is Ready and the editor holds what the parser last read.
        /// </summary>
        public ParserState CatchUp(bool force = false)
        {
            if (!IsEditorInDesignMode())
            {
                throw ApiException.NotInDesignMode();
            }

            var timing = Stopwatch.StartNew();
            var fingerprint = Capture();
            var why = force ? "a reparse was asked for" : WhyParseIsNeeded(fingerprint);
            if (why == null)
            {
                Logger.Info($"API catch-up: parser is Ready and nothing changed in {fingerprint.ModuleCount} modules; no parse requested (checked in {timing.ElapsedMilliseconds} ms).");
                return _state.Status;
            }

            Logger.Info($"API catch-up: parse requested because {why} (checked {fingerprint.ModuleCount} modules in {timing.ElapsedMilliseconds} ms).");
            var state = RequestParseAndWait();
            if (state == ParserState.Ready)
            {
                //The editor's list of modules may have been empty or stale before the run (the first one after a start), so read it again; but only trust what the run can have seen:
                //a module that changed or went away meanwhile means somebody edited during the run, and the next call must parse again.
                var after = Capture();
                _lastParsed = after.OnlyAddsTo(fingerprint) ? after : fingerprint;
                _lastParsedScheduled = Volatile.Read(ref _scheduled);
            }
            else
            {
                _lastParsed = null;
            }

            Logger.Info($"API catch-up: parser {state} after {timing.ElapsedMilliseconds} ms.");
            return state;
        }

        private ParserState RequestParseAndWait()
        {
            var changed = new ManualResetEventSlim(false);
            EventHandler<ParserStateEventArgs> onChange = (sender, args) => changed.Set();
            _state.StateChanged += onChange;
            try
            {
                var requested = _dispatcher.StartTask(() => _parseManager.OnParseRequested(this), CancellationToken.None);
                requested.Wait(TimeSpan.FromSeconds(30));

                //If nothing changed, the parser may stay Ready and never announce anything.
                changed.Wait(StartTimeout);

                var timer = Stopwatch.StartNew();
                while (!IsSettled(_state.Status))
                {
                    if (timer.Elapsed > ParseTimeout)
                    {
                        throw ApiException.Unavailable($"Rubberduck is still parsing ({_state.Status}). Try again in a moment.");
                    }

                    Thread.Sleep(100);
                }

                return _state.Status;
            }
            finally
            {
                _state.StateChanged -= onChange;
                changed.Dispose();
            }
        }

        private static bool IsSettled(ParserState state)
            => state == ParserState.Ready || state == ParserState.Error || state == ParserState.ResolverError || state == ParserState.UnexpectedError;
    }
}
