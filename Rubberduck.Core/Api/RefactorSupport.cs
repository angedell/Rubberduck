using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;
using Newtonsoft.Json.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.MoveToFolder;
using Rubberduck.VBEditor;

namespace Rubberduck.Api
{
    /// <summary>An MCP tool, described without the MCP endpoint's own private types, so that a refactoring can bring its tools with it.</summary>
    public class RefactorToolSpec
    {
        public string Name;
        public string Description;
        public JObject Properties;
        public string[] Required;
        public Func<JObject, string, object> Run;
    }

    /// <summary>What the refactoring routes and tools share: reading the request, finding the declaration it means, and running Rubberduck's own refactorings on the editor's thread.</summary>
    public static class RefactorArgs
    {
        public static JObject Body(string body)
        {
            if (string.IsNullOrWhiteSpace(body))
            {
                throw ApiException.BadRequest("A JSON body is required.");
            }

            return JToken.Parse(body) as JObject ?? throw ApiException.BadRequest("The body must be a JSON object.");
        }

        public static string Text(JObject body, string name)
        {
            var token = body[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            return token.Type == JTokenType.String ? token.Value<string>() : throw ApiException.BadRequest($"'{name}' must be a string.");
        }

        public static int? Number(JObject body, string name)
        {
            var token = body[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            return token.Type == JTokenType.Integer ? token.Value<int>() : throw ApiException.BadRequest($"'{name}' must be a whole number.");
        }

        public static bool Flag(JObject body, string name)
        {
            var token = body[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return false;
            }

            return token.Type == JTokenType.Boolean ? token.Value<bool>() : throw ApiException.BadRequest($"'{name}' must be true or false.");
        }

        public static JObject String(string description) => new JObject { ["type"] = "string", ["description"] = description };
        public static JObject Integer(string description) => new JObject { ["type"] = "integer", ["description"] = description };
        public static JObject Boolean(string description) => new JObject { ["type"] = "boolean", ["description"] = description };

        /// <summary>The 422 for a refactoring that Rubberduck, or the request, does not allow.</summary>
        public static ApiException Refused(string message) => new ApiException(422, "refactoring_refused", message);
    }

    /// <summary>A declaration that a request may mean, reduced to what the request can name it by. <see cref="Tag"/> is whatever the caller wants back.</summary>
    public class TargetCandidate
    {
        public string Name { get; set; }

        /// <summary>The declaration type as Rubberduck names it (Procedure, Variable, PropertyGet...).</summary>
        public string Kind { get; set; }

        /// <summary>Where the name is declared, 1-based.</summary>
        public int Line { get; set; }

        public int Column { get; set; }
        public int EndColumn { get; set; }
        public string Parent { get; set; }
        public object Tag { get; set; }
    }

    public static class RefactorTargets
    {
        /// <summary>
        /// Picks the declaration a request means: by <paramref name="name"/>, narrowed by <paramref name="kind"/> and by the <paramref name="line"/> (and column) where the name is declared.
        /// Several matches are an error that lists them, except the Get, Let and Set of one property, which are one thing to a refactoring.
        /// </summary>
        public static object Choose(IReadOnlyList<TargetCandidate> candidates, string module, string name, string kind, int? line, int? column)
        {
            if (string.IsNullOrWhiteSpace(name) && line == null)
            {
                throw ApiException.BadRequest("Say what to act on: 'name' (with 'line' if the name is not unique in the module), or 'line' and 'column'.");
            }

            if (line < 1 || column < 1)
            {
                throw ApiException.BadRequest("line and column are 1-based.");
            }

            if (column != null && line == null)
            {
                throw ApiException.BadRequest("'column' needs 'line'.");
            }

            var found = candidates
                .Where(c => string.IsNullOrWhiteSpace(name) || string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase))
                .Where(c => string.IsNullOrWhiteSpace(kind) || string.Equals(c.Kind, kind, StringComparison.OrdinalIgnoreCase))
                .Where(c => line == null || (c.Line == line && (column == null || (column >= c.Column && column <= c.EndColumn))))
                .ToList();

            if (found.Count == 0)
            {
                throw ApiException.NotFound($"'{module}' declares nothing that matches{(string.IsNullOrWhiteSpace(name) ? string.Empty : " '" + name + "'")}"
                                            + $"{(line == null ? string.Empty : " at line " + line)}. Look it up with the symbols route.");
            }

            if (found.Count > 1 && !(found.All(c => c.Kind.StartsWith("Property", StringComparison.Ordinal)) && found.Select(c => c.Name.ToUpperInvariant()).Distinct().Count() == 1))
            {
                var listed = string.Join("; ", found.Select(c => $"{c.Kind} {c.Name}{(c.Parent == null ? string.Empty : " in " + c.Parent)} at line {c.Line}"));
                throw ApiException.Conflict($"More than one declaration matches: {listed}. Add 'line' (and 'column') or 'kind' to say which.");
            }

            return found[0].Tag;
        }
    }

    /// <summary>
    /// Runs Rubberduck's refactorings for the API. They belong on the editor's thread (they rewrite code and rename components), one at a time, and a refactoring
    /// asks its presenter, a dialog for the user, for the details: <see cref="PresenterFactory"/> hands them a presenter that answers from the request instead.
    /// </summary>
    public class RefactorHost
    {
        private static readonly TimeSpan UiTimeout = TimeSpan.FromSeconds(120);
        private static readonly TimeSpan GateTimeout = TimeSpan.FromSeconds(5);

        private readonly IUiDispatcher _dispatcher;
        private readonly object _gate = new object();

        public RefactorHost(ParserSync parser, IUiDispatcher dispatcher)
        {
            Parser = parser;
            _dispatcher = dispatcher;
        }

        public ParserSync Parser { get; }

        /// <summary>Waits for the parser to catch up with the editor, so the declarations are those of the code as it is now. 409 if the editor is not in design mode.</summary>
        public string CatchUp(string project, out string projectId)
        {
            var (id, name) = Parser.ResolveProject(project);
            projectId = id;
            Parser.CatchUp();
            return name;
        }

        /// <summary>Runs <paramref name="work"/> on the editor's thread, alone: a second refactoring at the same time is refused with 409.</summary>
        public T Run<T>(Func<T> work)
        {
            if (!Monitor.TryEnter(_gate, GateTimeout))
            {
                throw ApiException.Conflict("Another refactoring is still running. Try again in a moment.");
            }

            try
            {
                //What goes wrong on the editor's thread is carried back as it is: a thread that ends a task with a cancellation would otherwise lose its type on the way.
                var task = _dispatcher.StartTask(() =>
                {
                    try
                    {
                        return (Value: work(), Failure: (Exception)null);
                    }
                    catch (Exception exception)
                    {
                        return (Value: default(T), Failure: exception);
                    }
                }, CancellationToken.None);

                if (!task.Wait(UiTimeout))
                {
                    throw ApiException.Unavailable("The editor did not respond in time. It may be showing a dialog or running code.");
                }

                var outcome = task.Result;
                if (outcome.Failure != null)
                {
                    throw Translate(outcome.Failure);
                }

                return outcome.Value;
            }
            finally
            {
                Monitor.Exit(_gate);
            }
        }

        /// <summary>What a refactoring throws, in terms of the API. What it does not know is returned as it is.</summary>
        public static Exception Translate(Exception failure)
        {
            switch (failure)
            {
                case ApiException _:
                    return failure;
                case RefactoringAbortedException _:
                    return RefactorArgs.Refused("Rubberduck stopped the refactoring.");
                case AffectedModuleIsStaleException stale:
                    return ApiException.Conflict($"'{stale.StaleModule.ComponentName}' was edited and not parsed yet. Ask for a reparse and retry.");
                case SuspendParserFailureException _:
                    return ApiException.Unavailable("Rubberduck is parsing right now, so the refactoring cannot be made. Try again in a moment.");
                case NoTargetFolderException _:
                    return ApiException.BadRequest("A target folder is required.");
                case InvalidTargetDeclarationException _:
                    return RefactorArgs.Refused("That declaration cannot be refactored this way (Rubberduck does not offer it for this kind of declaration, or it is not code of the project).");
                case RefactoringException _:
                    return RefactorArgs.Refused($"Rubberduck refused the refactoring ({failure.GetType().Name.Replace("Exception", string.Empty)}).");
                case COMException com:
                    return ApiException.Unavailable($"The editor refused the request ({com.Message}). It may be in break mode or running code.");
                default:
                    return failure;
            }
        }

        public static void Rethrow(Exception failure) => ExceptionDispatchInfo.Capture(failure).Throw();
    }

    /// <summary>
    /// Presenters for the refactorings that answer from the request instead of asking the user. Which one is made is registered per presenter type;
    /// what it does with the model is up to the caller, who sets it for the one refactoring in progress (see <see cref="RefactorHost.Run{T}"/>).
    /// </summary>
    public class PresenterFactory : IRefactoringPresenterFactory
    {
        private readonly Dictionary<Type, Func<object, object>> _creators = new Dictionary<Type, Func<object, object>>();

        public PresenterFactory Register<TPresenter, TModel>(Func<TModel, TPresenter> create)
            where TPresenter : class, IRefactoringPresenter<TModel>
            where TModel : class, IRefactoringModel
        {
            _creators[typeof(TPresenter)] = model => create((TModel)model);
            return this;
        }

        public TPresenter Create<TPresenter, TModel>(TModel model)
            where TPresenter : class, IRefactoringPresenter<TModel>
            where TModel : class, IRefactoringModel
            => _creators.TryGetValue(typeof(TPresenter), out var create) ? (TPresenter)create(model) : null;

        public void Release<TPresenter>(TPresenter presenter)
        {
        }
    }

    /// <summary>Finds the parser's declaration that a request names. The parser has to have caught up with the editor (<see cref="RefactorHost.CatchUp"/>).</summary>
    public static class RefactorDeclarations
    {
        /// <summary>
        /// The declaration in <paramref name="module"/> that the request names. <paramref name="include"/> narrows which declarations it may be (null: any but the project);
        /// <paramref name="viaReferences"/> also accepts a position that is a use of a declaration, as the definition route does.
        /// </summary>
        public static Declaration Find(ParserSync parser, string projectId, string module, string name, string kind, int? line, int? column,
            Func<Declaration, bool> include = null, bool viaReferences = true)
        {
            if (string.IsNullOrWhiteSpace(module))
            {
                throw ApiException.BadRequest("'module' is required.");
            }

            if (!parser.ModuleExists(projectId, module))
            {
                throw ApiException.NotFound($"The project has no module '{module}'.");
            }

            if (viaReferences && line != null && column != null && line >= 1 && column >= 1)
            {
                var point = new Selection(line.Value, column.Value);
                var used = parser.State.AllUserDeclarations
                    .Where(d => string.Equals(d.ProjectId, projectId, StringComparison.OrdinalIgnoreCase))
                    .SelectMany(d => d.References)
                    .FirstOrDefault(r => string.Equals(r.QualifiedModuleName.ComponentName, module, StringComparison.OrdinalIgnoreCase)
                                         && r.Selection.Contains(point)
                                         && (string.IsNullOrWhiteSpace(name) || string.Equals(r.IdentifierName, name, StringComparison.OrdinalIgnoreCase)));
                if (used?.Declaration != null && used.Declaration.IsUserDefined && (include == null || include(used.Declaration)))
                {
                    return used.Declaration;
                }
            }

            var candidates = parser.State.AllUserDeclarations
                .Where(d => string.Equals(d.ProjectId, projectId, StringComparison.OrdinalIgnoreCase)
                            && string.Equals(d.ComponentName, module, StringComparison.OrdinalIgnoreCase)
                            && d.DeclarationType != DeclarationType.Project
                            && (include == null || include(d)))
                .Select(d => new TargetCandidate
                {
                    Name = d.IdentifierName,
                    Kind = d.DeclarationType.ToString(),
                    Line = d.Selection.StartLine,
                    Column = d.Selection.StartColumn,
                    EndColumn = d.Selection.EndColumn,
                    Parent = d.ParentDeclaration?.IdentifierName,
                    Tag = d
                })
                .ToList();

            return (Declaration)RefactorTargets.Choose(candidates, module, name, kind, line, column);
        }
    }

    /// <summary>Stops a refactoring once its presenter has seen the model, so that nothing is changed: a preview.</summary>
    public sealed class RefactorPreviewCaptured : Exception
    {
    }
}
