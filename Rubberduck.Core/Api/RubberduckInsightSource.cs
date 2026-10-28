using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.CodeAnalysis.CodeMetrics;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.UIContext;
using Rubberduck.VBEditor.ComManagement;

namespace Rubberduck.Api
{
    /// <summary><see cref="IInsightSource"/> on the parser's state. Every call first waits for the parser to catch up with the editor.</summary>
    public class RubberduckInsightSource : IInsightSource
    {
        private readonly ParserSync _parser;
        private readonly ICodeMetricsAnalyst _metrics;
        private readonly IProjectsProvider _projects;
        private readonly IUiDispatcher _dispatcher;

        public RubberduckInsightSource(ParserSync parser, ICodeMetricsAnalyst metrics, IProjectsProvider projects, IUiDispatcher dispatcher)
        {
            _parser = parser;
            _metrics = metrics;
            _projects = projects;
            _dispatcher = dispatcher;
        }

        public ParserErrorsReport ParserErrors(string project)
        {
            var (id, name) = _parser.ResolveProject(project);
            var state = _parser.CatchUp();
            var report = new ParserErrorsReport { Project = name, ParserState = state.ToString() };
            foreach (var pair in _parser.State.ModuleExceptions.Where(e => SameProject(e.Item1.ProjectId, id)))
            {
                var exception = pair.Item2;
                report.Errors.Add(new ParserError
                {
                    Module = pair.Item1.ComponentName,
                    Line = exception.LineNumber,
                    Column = exception.Position,
                    Message = exception.Message,
                    Token = exception.OffendingSymbol?.Text,
                    CodeKind = exception.CodeKind.ToString()
                });
            }

            report.Errors = report.Errors.OrderBy(e => e.Module, StringComparer.OrdinalIgnoreCase).ThenBy(e => e.Line).ToList();
            return report;
        }

        public SymbolsReport Symbols(string project, string module, bool includeLocals, string kind)
        {
            var (id, name) = _parser.ResolveProject(project);
            _parser.CatchUp();
            var declarations = InModule(id, module).ToList();
            if (declarations.Count == 0 && !_parser.ModuleExists(id, module))
            {
                throw ApiException.NotFound($"The project has no module '{module}'.");
            }

            var symbols = declarations
                .Where(d => d.DeclarationType != DeclarationType.Project && !d.DeclarationType.HasFlag(DeclarationType.Module))
                .Where(d => includeLocals || !IsLocal(d))
                .Where(d => kind == null || string.Equals(d.DeclarationType.ToString(), kind, StringComparison.OrdinalIgnoreCase))
                .Select(ToSymbol)
                .OrderBy(s => s.StartLine)
                .ToList();

            return new SymbolsReport { Project = name, Module = module, Symbols = symbols };
        }

        public ReferencesReport References(string project, string module, string name, int limit)
        {
            var (id, projectName) = _parser.ResolveProject(project);
            _parser.CatchUp();

            var matches = _parser.State.AllUserDeclarations
                .Where(d => SameProject(d.ProjectId, id) && string.Equals(d.IdentifierName, name, StringComparison.OrdinalIgnoreCase))
                .Where(d => module == null || string.Equals(d.ComponentName, module, StringComparison.OrdinalIgnoreCase))
                .ToList();

            var report = new ReferencesReport { Project = projectName };
            foreach (var declaration in matches)
            {
                var references = declaration.References.OrderBy(r => r.QualifiedModuleName.ComponentName, StringComparer.OrdinalIgnoreCase).ThenBy(r => r.Selection.StartLine).ToList();
                report.Symbols.Add(new SymbolUsage
                {
                    Declaration = ToSymbol(declaration),
                    Total = references.Count,
                    Truncated = references.Count > limit,
                    References = references.Take(limit).Select(r => new SymbolReference
                    {
                        Module = r.QualifiedModuleName.ComponentName,
                        Line = r.Selection.StartLine,
                        Column = r.Selection.StartColumn,
                        EndColumn = r.Selection.EndColumn,
                        IsAssignment = r.IsAssignment,
                        In = r.ParentScoping?.IdentifierName
                    }).ToList()
                });
            }

            return report;
        }

        public DefinitionReport Definition(string project, string module, int line, int column)
        {
            var (id, _) = _parser.ResolveProject(project);
            _parser.CatchUp();
            var point = new Selection(line, column);

            var declarations = _parser.State.AllUserDeclarations.Where(d => SameProject(d.ProjectId, id)).ToList();
            var used = declarations.SelectMany(d => d.References)
                .FirstOrDefault(r => string.Equals(r.QualifiedModuleName.ComponentName, module, StringComparison.OrdinalIgnoreCase) && r.Selection.Contains(point));
            if (used != null && used.Declaration != null && used.Declaration.IsUserDefined)
            {
                return new DefinitionReport { Found = true, Identifier = used.IdentifierName, Declaration = ToSymbol(used.Declaration) };
            }

            //The point may be on the declaration itself.
            var declared = declarations.FirstOrDefault(d => string.Equals(d.ComponentName, module, StringComparison.OrdinalIgnoreCase)
                                                            && !d.DeclarationType.HasFlag(DeclarationType.Module) && d.Selection.Contains(point));
            return declared != null
                ? new DefinitionReport { Found = true, Identifier = declared.IdentifierName, Declaration = ToSymbol(declared) }
                : new DefinitionReport { Found = false, Identifier = used?.IdentifierName };
        }

        public IReadOnlyList<LibraryReference> Libraries(string project)
        {
            var (id, _) = _parser.ResolveProject(project);
            var task = _dispatcher.StartTask(() =>
            {
                var libraries = new List<LibraryReference>();
                var found = _projects.Project(id);
                using (var references = found.References)
                {
                    foreach (var reference in references)
                    {
                        using (reference)
                        {
                            libraries.Add(new LibraryReference
                            {
                                Name = reference.Name,
                                Guid = reference.Guid,
                                Description = reference.Description,
                                Major = reference.Major,
                                Minor = reference.Minor,
                                Path = reference.FullPath,
                                IsBuiltIn = reference.IsBuiltIn,
                                IsBroken = reference.IsBroken
                            });
                        }
                    }
                }

                return (IReadOnlyList<LibraryReference>)libraries;
            }, CancellationToken.None);

            if (!task.Wait(TimeSpan.FromSeconds(30)))
            {
                throw ApiException.Unavailable("The editor did not respond in time.");
            }

            return task.Result;
        }

        public IReadOnlyList<StructureFolder> Structure(string project)
        {
            var (id, _) = _parser.ResolveProject(project);
            _parser.CatchUp();
            return _parser.State.AllUserDeclarations
                .Where(d => SameProject(d.ProjectId, id) && d.DeclarationType.HasFlag(DeclarationType.Module))
                .GroupBy(d => d.CustomFolder ?? string.Empty)
                .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
                .Select(g => new StructureFolder { Folder = g.Key, Modules = g.Select(d => d.IdentifierName).OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList() })
                .ToList();
        }

        public IReadOnlyList<MetricValue> Metrics(string project, string module)
        {
            var (id, _) = _parser.ResolveProject(project);
            _parser.CatchUp();
            return _metrics.GetMetrics(_parser.State)
                .Where(r => SameProject(r.Declaration.ProjectId, id)
                            && (module == null || string.Equals(r.Declaration.ComponentName, module, StringComparison.OrdinalIgnoreCase)))
                .Select(r => new MetricValue
                {
                    Module = r.Declaration.ComponentName,
                    Member = r.Declaration.DeclarationType.HasFlag(DeclarationType.Module) ? string.Empty : r.Declaration.IdentifierName,
                    Metric = r.Metric.Name,
                    Value = r.Value
                })
                .OrderBy(m => m.Module, StringComparer.OrdinalIgnoreCase).ThenBy(m => m.Member, StringComparer.OrdinalIgnoreCase).ThenBy(m => m.Metric)
                .ToList();
        }

        private IEnumerable<Declaration> InModule(string projectId, string module)
            => _parser.State.AllUserDeclarations.Where(d => SameProject(d.ProjectId, projectId)
                                                            && string.Equals(d.ComponentName, module, StringComparison.OrdinalIgnoreCase));

        private static bool SameProject(string left, string right) => string.Equals(left, right, StringComparison.OrdinalIgnoreCase);

        //Parameters and variables inside a procedure; a module's own variables and constants are not local.
        private static bool IsLocal(Declaration declaration)
            => declaration.ParentDeclaration != null && !declaration.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Module)
                                                     && declaration.ParentDeclaration.DeclarationType != DeclarationType.Project
                                                     && (declaration.DeclarationType == DeclarationType.Parameter
                                                         || declaration.DeclarationType == DeclarationType.Variable
                                                         || declaration.DeclarationType == DeclarationType.Constant
                                                         || declaration.DeclarationType == DeclarationType.LineLabel);

        private static Symbol ToSymbol(Declaration declaration)
        {
            var span = declaration.Context != null ? declaration.Context.GetSelection() : declaration.Selection;
            return new Symbol
            {
                Module = declaration.ComponentName,
                Name = declaration.IdentifierName,
                Kind = declaration.DeclarationType.ToString(),
                Accessibility = declaration.Accessibility.ToString(),
                AsType = declaration.AsTypeName ?? string.Empty,
                Parent = declaration.ParentDeclaration?.IdentifierName,
                StartLine = span.StartLine,
                EndLine = span.EndLine,
                Annotations = declaration.Annotations.Select(a => a.Annotation.Name).ToList()
            };
        }
    }
}
