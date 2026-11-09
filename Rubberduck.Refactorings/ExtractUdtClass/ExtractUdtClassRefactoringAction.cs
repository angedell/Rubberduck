using System;
using System.Linq;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions.ExtractUdtClass;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.ExtractUdtClass
{
    /// <summary>Adds a class module with the properties of a user defined type, and with chainable methods if asked to.</summary>
    public class ExtractUdtClassRefactoringAction : RefactoringActionWithSuspension<ExtractUdtClassModel>
    {
        private readonly IAddComponentService _addComponentService;

        public ExtractUdtClassRefactoringAction(
            IParseManager parseManager,
            IRewritingManager rewritingManager,
            IAddComponentService addComponentService)
            : base(parseManager, rewritingManager)
        {
            _addComponentService = addComponentService;
        }

        // Adding the component makes the parser read the project again, so it has to be out of the way.
        protected override bool RequiresSuspension(ExtractUdtClassModel model) => true;

        protected override void Refactor(ExtractUdtClassModel model, IRewriteSession rewriteSession)
        {
            if (model.Clashes.Count > 0)
            {
                throw new MemberNamesClashException(model.Clashes);
            }

            var folderAnnotation = Declaration.GetModuleParent(model.Target)?.Annotations
                .FirstOrDefault(annotation => annotation.Annotation is FolderAnnotation);
            var folderAnnotationText = folderAnnotation != null ? $"'@{folderAnnotation.Context.GetText()}" : null;

            var body = ExtractUdtClassCodeBuilder.ClassBody(
                model.ClassName, model.Target.IdentifierName, model.Members, model.Style, folderAnnotationText);

            var projectId = model.Target.ProjectId;
            if (model.Style == ExtractUdtClassStyle.Fluent)
            {
                // The predeclared instance is an attribute, which only comes in with the module header.
                var code = $"{ExtractUdtClassCodeBuilder.ModuleHeader(model.ClassName)}{Environment.NewLine}{body}";
                _addComponentService.AddComponentWithAttributes(projectId, ComponentType.ClassModule, code, componentName: model.ClassName);
            }
            else
            {
                _addComponentService.AddComponent(projectId, ComponentType.ClassModule, body, componentName: model.ClassName);
            }
        }
    }
}
