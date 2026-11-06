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
    /// Warns about inconsistent implicit lower bounds of ParamArray arrays when 'Option Base 1' is specified.
    /// </summary>
    /// <why>
    /// The base of a ParamArray is always zero, regardless of any 'Option Base' setting.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Base 1 '<~ Implicit array lower bound is 1
    /// Public Sub DoSomething(ParamArray Values) '<<< inspection result here
    ///     Debug.Print LBound(Values) '<~ not 1
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
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
    internal sealed class InconsistentParamArrayBaseInspection : DeclarationInspectionBase
    {
        public InconsistentParamArrayBaseInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
        }

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            //The cheap test first: looking for 'Option Base' walks the whole module's parse tree, which is not worth doing for every declaration of the project.
            if (!(declaration is ParameterDeclaration parameter) || !parameter.IsParamArray)
            {
                return false;
            }

            return declaration.Context
                .GetAncestor<VBAParser.ModuleContext>()
                .GetDescendent<VBAParser.OptionBaseStmtContext>()?
                .numberLiteral()?.GetText() == "1";
        }

        protected override string ResultDescription(Declaration declaration)
        {
            // declaration is the ParamArray parameter
            return string.Format(InspectionResults.ResourceManager.GetString(nameof(InconsistentParamArrayBaseInspection), CultureInfo.CurrentUICulture), declaration.IdentifierName);
        }
    }
}
