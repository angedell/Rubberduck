using System;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Api
{
    /// <summary><see cref="IRenameSource"/> on Rubberduck's <see cref="RenameRefactoring"/>, with a presenter that answers from the request instead of showing the dialog.</summary>
    public class RubberduckRenameSource : IRenameSource
    {
        private readonly RefactorHost _host;
        private readonly RenameRefactoring _refactoring;

        //What the presenter does with the model of the refactoring in progress. Set and used inside RefactorHost.Run, which allows one at a time.
        private Func<RenameModel, RenameModel> _onShow;

        public RubberduckRenameSource(RefactorHost host, Func<IRefactoringPresenterFactory, RenameRefactoring> refactoring)
        {
            _host = host;
            _refactoring = refactoring(new PresenterFactory().Register<IRenamePresenter, RenameModel>(model => new ApiRenamePresenter(model, m => _onShow(m))));
        }

        /// <summary>Builds the API on what the add-in has registered; <paramref name="resolve"/> asks its container.</summary>
        public static RefactorRenameApi CreateApi(Func<Type, object> resolve)
        {
            var dispatcher = (IUiDispatcher)resolve(typeof(IUiDispatcher));
            var source = new RubberduckRenameSource(new RefactorHost((ParserSync)resolve(typeof(ParserSync)), dispatcher), presenters => new RenameRefactoring(
                (RenameRefactoringAction)resolve(typeof(RenameRefactoringAction)),
                new RefactoringUserInteraction<IRenamePresenter, RenameModel>(presenters, dispatcher),
                (IDeclarationFinderProvider)resolve(typeof(RubberduckParserState)),
                (IProjectsProvider)resolve(typeof(IProjectsProvider)),
                (ISelectionProvider)resolve(typeof(ISelectionProvider)),
                (ISelectedDeclarationProvider)resolve(typeof(ISelectedDeclarationProvider))));
            return new RefactorRenameApi(source, (ApiWriteApproval)resolve(typeof(ApiWriteApproval)));
        }

        public RenamePreview Preview(string project, RenameRequest request)
        {
            var name = _host.CatchUp(project, out var id);
            var target = RefactorDeclarations.Find(_host.Parser, id, request.Module, request.Name, request.Kind, request.Line, request.Column);
            RenamePreview preview = null;
            _host.Run(() =>
            {
                _onShow = model =>
                {
                    preview = Describe(name, model, request);
                    throw new RefactorPreviewCaptured();
                };

                try
                {
                    _refactoring.Refactor(target);
                }
                catch (RefactorPreviewCaptured)
                {
                }

                return true;
            });

            return preview ?? throw RefactorArgs.Refused("Rubberduck did not offer to rename that declaration.");
        }

        public RenameResult Apply(string project, RenameRequest request)
        {
            var name = _host.CatchUp(project, out var id);
            var target = RefactorDeclarations.Find(_host.Parser, id, request.Module, request.Name, request.Kind, request.Line, request.Column);
            RenameResult result = null;
            _host.Run(() =>
            {
                _onShow = model =>
                {
                    //Once more, now that nothing can change in between: what the refactoring will rename is what was approved.
                    var preview = Describe(name, model, request);
                    RefactorRenameApi.Verify(preview, request);
                    var renamed = model.Target;
                    result = new RenameResult
                    {
                        Project = name,
                        OldName = renamed.IdentifierName,
                        NewName = request.NewName,
                        Renamed = preview.ResolvedTarget ?? preview.Target,
                        ReferenceCount = preview.ReferenceCount,
                        ModulesAffected = preview.ModulesAffected,
                        Note = "Modules were rewritten; read again any you hold. The parser reparses them in the background."
                    };

                    model.NewName = request.NewName;
                    return model;
                };

                _refactoring.Refactor(target);
                return true;
            });

            return result;
        }

        private RenamePreview Describe(string project, RenameModel model, RenameRequest request)
        {
            var target = model.Target;
            var preview = new RenamePreview
            {
                Project = project,
                Target = ToTarget(model.InitialTarget),
                NewName = request.NewName
            };

            if (!model.Target.Equals(model.InitialTarget))
            {
                preview.ResolvedTarget = ToTarget(target);
                preview.ResolvedBecause = model.IsControlEventHandlerRename ? "control event handler"
                    : model.IsUserEventHandlerRename ? "user event"
                    : model.IsInterfaceMemberRename ? "interface member"
                    : "related declaration";
            }

            //What the dialog checks as the user types.
            preview.Problems.AddRange(VBAIdentifierValidator.SatisfiedInvalidIdentifierCriteria(request.NewName, target.DeclarationType));
            if (target.IdentifierName.Equals(request.NewName, StringComparison.OrdinalIgnoreCase))
            {
                preview.Problems.Add(target.IdentifierName.Equals(request.NewName)
                    ? "It is the current name."
                    : "Only the case differs, which Rubberduck does not rename.");
            }

            preview.Valid = preview.Problems.Count == 0;
            if (preview.Valid)
            {
                preview.Conflicts.AddRange(_host.Parser.State.DeclarationFinder.FindNewDeclarationNameConflicts(request.NewName, target)
                    .Select(d => d.QualifiedName.ToString()));
            }

            preview.ReferenceCount = target.References.Count();
            preview.ModulesAffected = target.References.Select(r => r.QualifiedModuleName.ComponentName)
                .Concat(new[] { target.ComponentName })
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(m => m, StringComparer.OrdinalIgnoreCase)
                .ToList();
            return preview;
        }

        private static RenameTarget ToTarget(Declaration declaration) => new RenameTarget
        {
            Module = declaration.ComponentName,
            Name = declaration.IdentifierName,
            Kind = declaration.DeclarationType.ToString(),
            Line = declaration.Selection.StartLine
        };

        private class ApiRenamePresenter : IRenamePresenter
        {
            private readonly Func<RenameModel, RenameModel> _show;

            public ApiRenamePresenter(RenameModel model, Func<RenameModel, RenameModel> show)
            {
                Model = model;
                _show = show;
            }

            public RenameModel Model { get; }

            public RenameModel Show() => _show(Model);

            public RenameModel Show(Declaration target)
            {
                Model.Target = target;
                return Show();
            }
        }
    }
}
