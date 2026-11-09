using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Refactorings.ExtractUdtClass;

namespace Rubberduck.Api
{
    /// <summary>
    /// <see cref="IExtractUdtClassSource"/> on Rubberduck's <see cref="ExtractUdtClassRefactoring"/>: the model the dialog would start from, with the
    /// dialog's two questions (name and style) answered from the request, and then the refactoring's action.
    /// </summary>
    public class RubberduckExtractUdtClassSource : IExtractUdtClassSource
    {
        private readonly RefactorHost _host;
        private readonly ExtractUdtClassRefactoring _refactoring;
        private readonly ExtractUdtClassRefactoringAction _action;

        public RubberduckExtractUdtClassSource(RefactorHost host, ExtractUdtClassRefactoring refactoring, ExtractUdtClassRefactoringAction action)
        {
            _host = host;
            _refactoring = refactoring;
            _action = action;
        }

        /// <summary>Builds the API on what the add-in has registered; <paramref name="resolve"/> asks its container.</summary>
        public static RefactorExtractUdtClassApi CreateApi(Func<Type, object> resolve)
        {
            var host = new RefactorHost((ParserSync)resolve(typeof(ParserSync)), (IUiDispatcher)resolve(typeof(IUiDispatcher)));
            return new RefactorExtractUdtClassApi(
                new RubberduckExtractUdtClassSource(
                    host,
                    (ExtractUdtClassRefactoring)resolve(typeof(ExtractUdtClassRefactoring)),
                    (ExtractUdtClassRefactoringAction)resolve(typeof(ExtractUdtClassRefactoringAction))),
                (ApiWriteApproval)resolve(typeof(ApiWriteApproval)));
        }

        private Declaration Find(string projectId, ExtractUdtClassRequest request)
            => RefactorDeclarations.Find(_host.Parser, projectId, request.Module, request.Name, "UserDefinedType", request.Line, request.Column,
                d => d.DeclarationType == DeclarationType.UserDefinedType, viaReferences: false);

        private ExtractUdtClassModel Model(Declaration target, ExtractUdtClassRequest request)
        {
            var model = _refactoring.CreateModel(target);
            if (!string.IsNullOrWhiteSpace(request.ClassName))
            {
                model.ClassName = request.ClassName;
            }

            model.Style = string.Equals(request.Style, ExtractUdtClassRequest.Fluent, StringComparison.OrdinalIgnoreCase)
                ? ExtractUdtClassStyle.Fluent
                : ExtractUdtClassStyle.Plain;

            foreach (var member in model.Members)
            {
                member.IsRequired = request.Required.Any(name => string.Equals(name, member.Name, StringComparison.OrdinalIgnoreCase));
            }

            return model;
        }

        private static string StyleName(ExtractUdtClassModel model)
            => model.Style == ExtractUdtClassStyle.Fluent ? ExtractUdtClassRequest.Fluent : ExtractUdtClassRequest.Plain;

        private static ExtractUdtClassPreview Describe(string project, ExtractUdtClassModel model, ExtractUdtClassRequest request)
        {
            var preview = new ExtractUdtClassPreview
            {
                Project = project,
                Module = model.Target.ComponentName,
                Type = model.Target.IdentifierName,
                ClassName = model.ClassName,
                Style = StyleName(model),
                Members = model.Members.Select(m => new ExtractedMember
                {
                    Name = m.Name,
                    Type = m.IsArray ? m.TypeName + "()" : m.TypeName,
                    Access = m.IsUserDefinedType ? "left out" : m.IsArray ? "readOnly" : "readWrite",
                    Required = m.IsRequired && m.CanBeRequired && model.Style == ExtractUdtClassStyle.Fluent
                }).ToList(),
                Code = ExtractUdtClassCodeBuilder.ClassBody(model.ClassName, model.Target.IdentifierName, model.Members, model.Style),
                Problems = model.Problems.ToList()
            };

            preview.Problems.AddRange(RequiredProblems(model, request));
            return preview;
        }

        //What the dialog would not let the user do: ask for a Create argument that is not there, cannot be one, or has no Create to go to.
        private static IEnumerable<string> RequiredProblems(ExtractUdtClassModel model, ExtractUdtClassRequest request)
        {
            if (request.Required.Count > 0 && model.Style != ExtractUdtClassStyle.Fluent)
            {
                yield return "'required' members are the arguments of Create, which only the fluent style has.";
            }

            foreach (var name in request.Required)
            {
                var member = model.Members.FirstOrDefault(m => string.Equals(m.Name, name, StringComparison.OrdinalIgnoreCase));
                if (member == null)
                {
                    yield return $"'{name}' is not a member of the type '{model.Target.IdentifierName}'.";
                }
                else if (!member.CanBeRequired)
                {
                    yield return $"'{member.Name}' cannot be required: an array or a Type inside the Type cannot be an argument of Create.";
                }
            }
        }

        public ExtractUdtClassPreview Preview(string project, ExtractUdtClassRequest request)
        {
            var name = _host.CatchUp(project, out var id);
            return Describe(name, Model(Find(id, request), request), request);
        }

        public ExtractUdtClassResult Apply(string project, ExtractUdtClassRequest request)
        {
            var name = _host.CatchUp(project, out var id);
            var target = Find(id, request);
            return _host.Run(() =>
            {
                //Once more, now that nothing can change in between.
                var model = Model(target, request);
                RefactorExtractUdtClassApi.Verify(Describe(name, model, request));
                _action.Refactor(model);
                return new ExtractUdtClassResult
                {
                    Project = name,
                    Type = model.Target.IdentifierName,
                    ClassName = model.ClassName,
                    Style = StyleName(model),
                    Note = "The class module was added; list the modules again. The parser reparses the project in the background."
                };
            });
        }
    }
}
