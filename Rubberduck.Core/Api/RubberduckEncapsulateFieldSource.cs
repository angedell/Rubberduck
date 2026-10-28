using System;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Api
{
    /// <summary>
    /// <see cref="IEncapsulateFieldSource"/> on Rubberduck's <see cref="EncapsulateFieldRefactoring"/>, with a presenter that sets the model from the request
    /// instead of showing the dialog, the way the dialog's view model does: strategy first, then the property name and read-only flag of the field.
    /// </summary>
    public class RubberduckEncapsulateFieldSource : IEncapsulateFieldSource
    {
        private readonly RefactorHost _host;
        private readonly EncapsulateFieldRefactoring _refactoring;

        //What the presenter does with the model of the refactoring in progress. Set and used inside RefactorHost.Run, which allows one at a time.
        private Func<EncapsulateFieldModel, EncapsulateFieldModel> _onShow;

        public RubberduckEncapsulateFieldSource(RefactorHost host, Func<IRefactoringPresenterFactory, EncapsulateFieldRefactoring> refactoring)
        {
            _host = host;
            _refactoring = refactoring(new PresenterFactory().Register<IEncapsulateFieldPresenter, EncapsulateFieldModel>(model => new ApiEncapsulateFieldPresenter(model, m => _onShow(m))));
        }

        /// <summary>Builds the API on what the add-in has registered; <paramref name="resolve"/> asks its container.</summary>
        public static RefactorEncapsulateFieldApi CreateApi(Func<Type, object> resolve)
        {
            var dispatcher = (IUiDispatcher)resolve(typeof(IUiDispatcher));
            var source = new RubberduckEncapsulateFieldSource(new RefactorHost((ParserSync)resolve(typeof(ParserSync)), dispatcher), presenters => new EncapsulateFieldRefactoring(
                (EncapsulateFieldRefactoringAction)resolve(typeof(EncapsulateFieldRefactoringAction)),
                (EncapsulateFieldPreviewProvider)resolve(typeof(EncapsulateFieldPreviewProvider)),
                (IEncapsulateFieldModelFactory)resolve(typeof(IEncapsulateFieldModelFactory)),
                new RefactoringUserInteraction<IEncapsulateFieldPresenter, EncapsulateFieldModel>(presenters, dispatcher),
                (ISelectionProvider)resolve(typeof(ISelectionProvider)),
                (ISelectedDeclarationProvider)resolve(typeof(ISelectedDeclarationProvider))));
            return new RefactorEncapsulateFieldApi(source, (ApiWriteApproval)resolve(typeof(ApiWriteApproval)));
        }

        //The refactoring takes a field of the module, not a variable inside a procedure (its own menu command only offers the former).
        private Declaration Find(string projectId, EncapsulateFieldRequest request)
            => RefactorDeclarations.Find(_host.Parser, projectId, request.Module, request.Name, "Variable", request.Line, request.Column,
                d => d.DeclarationType == DeclarationType.Variable && d.ParentDeclaration != null && d.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Module));

        public EncapsulateFieldPreview Preview(string project, EncapsulateFieldRequest request)
        {
            var name = _host.CatchUp(project, out var id);
            var target = Find(id, request);
            EncapsulateFieldPreview preview = null;
            _host.Run(() =>
            {
                _onShow = model =>
                {
                    preview = Describe(name, model, target, request);
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

            return preview ?? throw RefactorArgs.Refused("Rubberduck did not offer to encapsulate that field.");
        }

        public EncapsulateFieldResult Apply(string project, EncapsulateFieldRequest request)
        {
            var name = _host.CatchUp(project, out var id);
            var target = Find(id, request);
            EncapsulateFieldResult result = null;
            _host.Run(() =>
            {
                _onShow = model =>
                {
                    //Once more, now that nothing can change in between.
                    var preview = Describe(name, model, target, request);
                    RefactorEncapsulateFieldApi.Verify(preview);
                    result = new EncapsulateFieldResult
                    {
                        Project = name,
                        Module = preview.Module,
                        Field = preview.Field,
                        Strategy = preview.Strategy,
                        Properties = preview.Properties,
                        Note = "The module was rewritten; read it again. The parser reparses it in the background."
                    };

                    return model;
                };

                _refactoring.Refactor(target);
                return true;
            });

            return result;
        }

        private static EncapsulateFieldPreview Describe(string project, EncapsulateFieldModel model, Declaration target, EncapsulateFieldRequest request)
        {
            var preview = new EncapsulateFieldPreview { Project = project, Module = target.ComponentName, Field = target.IdentifierName, Line = target.Selection.StartLine };

            if (request.Strategy != null)
            {
                model.EncapsulateFieldStrategy = request.Strategy.Equals(EncapsulateFieldRequest.UdtMember, StringComparison.OrdinalIgnoreCase)
                    ? EncapsulateFieldStrategy.ConvertFieldsToUDTMembers
                    : EncapsulateFieldStrategy.UseBackingFields;
            }

            preview.Strategy = model.EncapsulateFieldStrategy == EncapsulateFieldStrategy.ConvertFieldsToUDTMembers ? EncapsulateFieldRequest.UdtMember : EncapsulateFieldRequest.BackingField;

            var candidate = model.SelectedFieldCandidates.FirstOrDefault(c => c.Declaration.Equals(target)) ?? model.SelectedFieldCandidates.FirstOrDefault();
            if (candidate == null)
            {
                preview.Problems.Add("Rubberduck does not offer to encapsulate this field.");
                return preview;
            }

            if (request.PropertyName != null)
            {
                if (candidate is IUserDefinedTypeCandidate udt && udt.TypeDeclarationIsPrivate)
                {
                    preview.Problems.Add("The properties of a field of a private Type are named after the Type's members; leave out 'propertyName'.");
                }
                else
                {
                    candidate.PropertyIdentifier = request.PropertyName;
                }
            }

            if (request.ReadOnly.HasValue)
            {
                candidate.IsReadOnly = request.ReadOnly.Value;
                //The candidate has the last word: an array cannot be written through a property, a field that other modules assign cannot be read-only.
                if (candidate.IsReadOnly != request.ReadOnly.Value)
                {
                    preview.Problems.Add(request.ReadOnly.Value
                        ? "Other modules assign the field, so its property cannot be read-only."
                        : "An array cannot be given a Property Let/Set, only a Property Get.");
                }
            }

            //What the dialog shows beside the list: each selected field's validation message.
            foreach (var selected in model.SelectedFieldCandidates)
            {
                if (!selected.TryValidateEncapsulationAttributes(out var problem) && !string.IsNullOrWhiteSpace(problem))
                {
                    preview.Problems.Add(problem);
                }

                preview.Properties.Add(new EncapsulatedProperty
                {
                    Field = selected.TargetID,
                    PropertyName = selected.PropertyIdentifier,
                    BackingField = selected.BackingIdentifier,
                    ReadOnly = selected.IsReadOnly,
                    CanBeReadWrite = selected.CanBeReadWrite,
                    AssignedExternally = selected.IsAssignedExternally
                });
            }

            preview.Valid = preview.Problems.Count == 0;
            if (preview.Valid)
            {
                preview.Code = model.PreviewRefactoring();
            }

            return preview;
        }

        private class ApiEncapsulateFieldPresenter : IEncapsulateFieldPresenter
        {
            private readonly EncapsulateFieldModel _model;
            private readonly Func<EncapsulateFieldModel, EncapsulateFieldModel> _show;

            public ApiEncapsulateFieldPresenter(EncapsulateFieldModel model, Func<EncapsulateFieldModel, EncapsulateFieldModel> show)
            {
                _model = model;
                _show = show;
            }

            public EncapsulateFieldModel Show() => _show(_model);
        }
    }
}
