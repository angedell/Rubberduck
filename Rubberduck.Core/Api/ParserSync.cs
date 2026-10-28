using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
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
            //What counts is a module being scheduled for a parse (new, edited, marked as modified).
            _state.ModuleStateChanged += (sender, args) =>
            {
                if (args.State == ParserState.Pending)
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

        /// <summary>Asks for a parse, waits until the parser is done, and says how it ended.</summary>
        public ParserState CatchUp()
        {
            if (!IsEditorInDesignMode())
            {
                throw ApiException.NotInDesignMode();
            }

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
