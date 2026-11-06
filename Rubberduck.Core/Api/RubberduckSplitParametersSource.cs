using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.SplitParameters;

namespace Rubberduck.Api
{
    /// <summary><see cref="ISplitParametersSource"/> on Rubberduck's <see cref="SplitParametersRefactoring"/>, which needs no dialog.</summary>
    public class RubberduckSplitParametersSource : ISplitParametersSource
    {
        private readonly RefactorHost _host;
        private readonly SplitParametersRefactoring _refactoring;

        public RubberduckSplitParametersSource(RefactorHost host, SplitParametersRefactoring refactoring)
        {
            _host = host;
            _refactoring = refactoring;
        }

        /// <summary>Builds the API on what the add-in has registered; <paramref name="resolve"/> asks its container.</summary>
        public static RefactorSplitParametersApi CreateApi(Func<Type, object> resolve)
        {
            var host = new RefactorHost((ParserSync)resolve(typeof(ParserSync)), (IUiDispatcher)resolve(typeof(IUiDispatcher)));
            return new RefactorSplitParametersApi(
                new RubberduckSplitParametersSource(host, (SplitParametersRefactoring)resolve(typeof(SplitParametersRefactoring))),
                (ApiWriteApproval)resolve(typeof(ApiWriteApproval)));
        }

        private List<Declaration> Members(string projectId, SplitParametersRequest request)
        {
            if (request.Module != null && !_host.Parser.ModuleExists(projectId, request.Module))
            {
                throw ApiException.NotFound($"The project has no module '{request.Module}'.");
            }

            return _refactoring.MembersToSplit(projectId)
                .Where(m => request.Module == null || string.Equals(m.QualifiedModuleName.ComponentName, request.Module, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        private List<SplitMember> Describe(IEnumerable<Declaration> members) => members.Select(m => new SplitMember
        {
            Module = m.QualifiedModuleName.ComponentName,
            Member = m.IdentifierName,
            Parameters = _refactoring.ParameterCount(m)
        }).ToList();

        public SplitParametersPreview Preview(string project, SplitParametersRequest request)
        {
            var name = _host.CatchUp(project, out var id);
            return new SplitParametersPreview { Project = name, Members = Describe(Members(id, request)) };
        }

        public SplitParametersResult Apply(string project, SplitParametersRequest request)
        {
            var name = _host.CatchUp(project, out var id);
            return _host.Run(() =>
            {
                //Looked at again, now that nothing can change in between.
                var members = Members(id, request);
                _refactoring.Refactor(members);
                return new SplitParametersResult
                {
                    Project = name,
                    Split = Describe(members),
                    Note = "The parameter lists were rewritten; read the modules again. The parser reparses them in the background."
                };
            });
        }
    }
}
