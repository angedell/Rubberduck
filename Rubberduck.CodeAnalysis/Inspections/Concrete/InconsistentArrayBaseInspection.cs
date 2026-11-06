using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using System.Globalization;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about inconsistent implicit lower bounds of VBA.Array arrays when 'Option Base 1' is specified.
    /// </summary>
    /// <why>
    /// The base of an array obtained from a qualified 'VBA.Array' function call is always zero, regardless of any 'Option Base' setting.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Base 1 '<~ Implicit array lower bound is 1
    /// 
    /// Public Sub DoSomething()
    ///     Dim Values As Variant
    ///     
    ///     Values = Array(42)
    ///     Debug.Print LBound(Values) '<~ 1 as per Option Base
    ///     
    ///     Values = VBA.Array(42) '<<< inspection result here
    ///     Debug.Print LBound(Values) '<~ not 1
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// 'implicit: Option Base 0
    /// 
    /// Public Sub DoSomething()
    ///     Dim Values As Variant
    ///     
    ///     Values = Array(42)
    ///     Debug.Print LBound(Values) '<~ 0
    ///     
    ///     Values = VBA.Array(42)
    ///     Debug.Print LBound(Values) '<~ also 0
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal class InconsistentArrayBaseInspection : IdentifierReferenceInspectionBase
    {
        public InconsistentArrayBaseInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            //The cheap tests first: looking for 'Option Base' walks the whole module's parse tree, which is not worth doing for every reference of the project.
            if (reference.Declaration.ProjectName != "VBA" || reference.Declaration.IdentifierName != "Array"
                || reference.QualifyingReference?.Declaration.IdentifierName != "VBA")
            {
                return false;
            }

            return reference.Context
                .GetAncestor<VBAParser.ModuleContext>()
                .GetDescendent<VBAParser.OptionBaseStmtContext>()?
                .numberLiteral()?.GetText() == "1";
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            // reference.Declaration is the VBA.Array function
            // we could inspect the context to find a possible LHS variable being assigned, but VBA.Array could also be an argument
            // so it's not a given that there's a relevant identifier to call out, so the resource string does not have any placeholders.
            return InspectionResults.ResourceManager.GetString(nameof(InconsistentArrayBaseInspection), CultureInfo.CurrentUICulture);
        }
    }
}
