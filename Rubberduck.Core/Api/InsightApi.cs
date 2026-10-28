using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Api
{
    /// <summary>What Rubberduck knows about the code that a plain read of it does not tell: parse errors, declarations, who uses what, libraries, structure and metrics.</summary>
    public interface IInsightSource
    {
        ParserErrorsReport ParserErrors(string project);
        SymbolsReport Symbols(string project, string module, bool includeLocals, string kind);
        ReferencesReport References(string project, string module, string name, int limit);
        DefinitionReport Definition(string project, string module, int line, int column);
        IReadOnlyList<LibraryReference> Libraries(string project);
        IReadOnlyList<StructureFolder> Structure(string project);
        IReadOnlyList<MetricValue> Metrics(string project, string module);
    }

    /// <summary>The read-only questions agents can ask about a project. Checks what the caller sent, caps what comes back, asks nobody for approval.</summary>
    public class InsightApi
    {
        public const int DefaultReferenceLimit = 200;
        public const int MaxReferenceLimit = 2000;
        public const int MaxSymbols = 2000;

        private readonly IInsightSource _source;

        public InsightApi(IInsightSource source)
        {
            _source = source;
        }

        public ParserErrorsReport ParserErrors(string project) => _source.ParserErrors(project);

        public SymbolsReport Symbols(string project, string module, bool includeLocals, string kind)
        {
            if (string.IsNullOrWhiteSpace(module))
            {
                throw ApiException.BadRequest("'module' is required.");
            }

            var report = _source.Symbols(project, module, includeLocals, string.IsNullOrWhiteSpace(kind) ? null : kind);
            report.Total = report.Symbols.Count;
            report.Truncated = report.Symbols.Count > MaxSymbols;
            report.Symbols = report.Symbols.Take(MaxSymbols).ToList();
            return report;
        }

        public ReferencesReport References(string project, string module, string name, int? limit)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw ApiException.BadRequest("'name' is required.");
            }

            if (limit < 1 || limit > MaxReferenceLimit)
            {
                throw ApiException.BadRequest($"limit must be between 1 and {MaxReferenceLimit}.");
            }

            return _source.References(project, string.IsNullOrWhiteSpace(module) ? null : module, name, limit ?? DefaultReferenceLimit);
        }

        public DefinitionReport Definition(string project, string module, int line, int column)
        {
            if (string.IsNullOrWhiteSpace(module))
            {
                throw ApiException.BadRequest("'module' is required.");
            }

            if (line < 1 || column < 1)
            {
                throw ApiException.BadRequest("line and column are 1-based.");
            }

            return _source.Definition(project, module, line, column);
        }

        public IReadOnlyList<LibraryReference> Libraries(string project) => _source.Libraries(project);

        public IReadOnlyList<StructureFolder> Structure(string project) => _source.Structure(project);

        public IReadOnlyList<MetricValue> Metrics(string project, string module)
            => _source.Metrics(project, string.IsNullOrWhiteSpace(module) ? null : module);
    }
}
