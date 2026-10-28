using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Api
{
    /// <summary>
    /// <see cref="IMoveToFolderSource"/> on Rubberduck's <see cref="MoveToFolderRefactoring"/>, with a presenter that sets the folder from the request instead of showing
    /// the dialog. The refactoring takes one module as its target; the presenter adds the others to the model, as the dialog does for a selection of several.
    /// </summary>
    public class RubberduckMoveToFolderSource : IMoveToFolderSource
    {
        private readonly RefactorHost _host;
        private readonly MoveToFolderRefactoring _refactoring;

        //What the presenter does with the model of the refactoring in progress. Set and used inside RefactorHost.Run, which allows one at a time.
        private Func<MoveMultipleToFolderModel, MoveMultipleToFolderModel> _onShow;

        public RubberduckMoveToFolderSource(RefactorHost host, Func<IRefactoringPresenterFactory, MoveToFolderRefactoring> refactoring)
        {
            _host = host;
            _refactoring = refactoring(new PresenterFactory().Register<IMoveMultipleToFolderPresenter, MoveMultipleToFolderModel>(model => new ApiMoveToFolderPresenter(model, m => _onShow(m))));
        }

        /// <summary>Builds the API on what the add-in has registered; <paramref name="resolve"/> asks its container.</summary>
        public static RefactorMoveToFolderApi CreateApi(Func<Type, object> resolve)
        {
            var dispatcher = (IUiDispatcher)resolve(typeof(IUiDispatcher));
            var source = new RubberduckMoveToFolderSource(new RefactorHost((ParserSync)resolve(typeof(ParserSync)), dispatcher), presenters => new MoveToFolderRefactoring(
                (MoveMultipleToFolderRefactoringAction)resolve(typeof(MoveMultipleToFolderRefactoringAction)),
                (ISelectedDeclarationProvider)resolve(typeof(ISelectedDeclarationProvider)),
                (ISelectionProvider)resolve(typeof(ISelectionProvider)),
                new RefactoringUserInteraction<IMoveMultipleToFolderPresenter, MoveMultipleToFolderModel>(presenters, dispatcher),
                (RubberduckParserState)resolve(typeof(RubberduckParserState))));
            return new RefactorMoveToFolderApi(source, (ApiWriteApproval)resolve(typeof(ApiWriteApproval)));
        }

        private List<ModuleDeclaration> Find(string projectId, MoveToFolderRequest request)
        {
            var found = new List<ModuleDeclaration>();
            foreach (var module in request.Modules)
            {
                if (!_host.Parser.ModuleExists(projectId, module))
                {
                    throw ApiException.NotFound($"The project has no module '{module}'.");
                }

                found.Add(_host.Parser.State.AllUserDeclarations.OfType<ModuleDeclaration>()
                    .FirstOrDefault(d => string.Equals(d.ProjectId, projectId, StringComparison.OrdinalIgnoreCase)
                                         && string.Equals(d.ComponentName, module, StringComparison.OrdinalIgnoreCase))
                    ?? throw RefactorArgs.Refused($"'{module}' has no declaration to move; it may have failed to parse."));
            }

            return found;
        }

        public MoveToFolderPreview Preview(string project, MoveToFolderRequest request)
        {
            var name = _host.CatchUp(project, out var id);
            var targets = Find(id, request);
            MoveToFolderPreview preview = null;
            _host.Run(() =>
            {
                _onShow = model =>
                {
                    preview = Describe(name, model, targets, request);
                    throw new RefactorPreviewCaptured();
                };

                try
                {
                    _refactoring.Refactor(targets[0]);
                }
                catch (RefactorPreviewCaptured)
                {
                }

                return true;
            });

            return preview ?? throw RefactorArgs.Refused("Rubberduck did not offer to move that module.");
        }

        public MoveToFolderResult Apply(string project, MoveToFolderRequest request)
        {
            var name = _host.CatchUp(project, out var id);
            var targets = Find(id, request);
            MoveToFolderResult result = null;
            _host.Run(() =>
            {
                _onShow = model =>
                {
                    //Once more, now that nothing can change in between.
                    var preview = Describe(name, model, targets, request);
                    RefactorMoveToFolderApi.Verify(preview);
                    result = new MoveToFolderResult
                    {
                        Project = name,
                        Folder = request.Folder,
                        Moved = preview.Moves.Where(m => m.Changes).Select(m => m.Module).ToList(),
                        Unchanged = preview.Moves.Where(m => !m.Changes).Select(m => m.Module).ToList(),
                        Note = "The modules' '@Folder annotations were written; read them again. The parser reparses them in the background."
                    };

                    return model;
                };

                _refactoring.Refactor(targets[0]);
                return true;
            });

            return result;
        }

        private static MoveToFolderPreview Describe(string project, MoveMultipleToFolderModel model, List<ModuleDeclaration> targets, MoveToFolderRequest request)
        {
            //The refactoring started with the first module; the dialog would let the user add more.
            foreach (var target in targets.Where(t => !model.Targets.Contains(t)))
            {
                model.Targets.Add(target);
            }

            model.TargetFolder = request.Folder;

            var problems = RefactorMoveToFolderApi.FolderProblems(request.Folder);
            return new MoveToFolderPreview
            {
                Project = project,
                Folder = request.Folder,
                Valid = problems.Count == 0,
                Problems = problems,
                Moves = model.Targets.Select(t => new ModuleMove
                {
                    Module = t.ComponentName,
                    FromFolder = t.CustomFolder,
                    //Without an annotation the module is only "in" the project's folder by default; the refactoring writes one.
                    Changes = !string.Equals(t.CustomFolder, request.Folder, StringComparison.Ordinal) || !t.Annotations.Any(a => a.Annotation is FolderAnnotation)
                }).ToList()
            };
        }

        private class ApiMoveToFolderPresenter : IMoveMultipleToFolderPresenter
        {
            private readonly MoveMultipleToFolderModel _model;
            private readonly Func<MoveMultipleToFolderModel, MoveMultipleToFolderModel> _show;

            public ApiMoveToFolderPresenter(MoveMultipleToFolderModel model, Func<MoveMultipleToFolderModel, MoveMultipleToFolderModel> show)
            {
                _model = model;
                _show = show;
            }

            public MoveMultipleToFolderModel Show() => _show(_model);
        }
    }
}
