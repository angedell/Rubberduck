using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Api;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Ship
{
    /// <summary>
    /// The analysis stage of tree shaking (ADR 0004) over the parser state of the open project: roots, then reachability over the resolved references,
    /// then what is left over, by name. It changes nothing, and it does not guess: if the parser is not ready, or any module did not parse or resolve, it says so and returns null.
    /// </summary>
    /// <remarks>
    /// Needs a thread that is not the editor's: <see cref="ParserSync.CatchUp"/> asks the editor's thread to parse and waits for it.
    /// A unit is a procedure group (Get/Let/Set of a name), a module-level variable statement or constant statement (all names of one statement stand or fall together),
    /// a Type, an Enum, a Declare, or a module. A unit is kept when it is a root or when a kept unit references it. The references of a kept unit are those inside its own span;
    /// references outside any unit (Implements, module-level statements) belong to the module and count once the module is kept.
    /// </remarks>
    public class ParserStateShipAnalysis : IShipAnalysis
    {
        private static readonly string[] DocumentHandlerPrefixes = { "Chart_", "Worksheet_", "Workbook_", "Document_", "Application_", "Session_" };
        private static readonly string[] AutoMacros = { "Auto_Open", "Auto_Close", "Auto_Add", "Auto_Remove", "Auto_Activate", "Auto_Deactivate" };

        private readonly RubberduckParserState _state;
        private readonly IStringParser _parser;
        private readonly Func<ParserState> _catchUp;
        private readonly Func<int> _generation;

        /// <summary>The analysis as Rubberduck runs it: the parser is asked to catch up with the editor and waited for.</summary>
        public ParserStateShipAnalysis(ParserSync sync, IStringParser parser) : this(sync.State, parser, () => sync.CatchUp(), () => sync.Generation)
        {
        }

        /// <remarks>Not public so that the container does not try to build it.</remarks>
        /// <param name="catchUp">Brings the parser up to date and returns where it ended. A test that parsed synchronously passes the state's status.</param>
        /// <param name="generation">Changes when the parser state may have changed (<see cref="ParserSync.Generation"/>). Null: nothing is checked.</param>
        internal ParserStateShipAnalysis(RubberduckParserState state, IStringParser parser, Func<ParserState> catchUp, Func<int> generation = null)
        {
            _state = state;
            _parser = parser;
            _catchUp = catchUp;
            _generation = generation ?? (() => 0);
        }

        #region Model

        private class ModuleInfo
        {
            public string Name;
            public ComponentType Type;
            public Declaration Declaration;
            public Unit Self;
            public List<Unit> Units = new List<Unit>();
            public List<ModuleInfo> Implements = new List<ModuleInfo>();
            /// <summary>Interfaces from a type library (IRibbonControl) that the class implements: they are not in the project, so only the names of the members are known.</summary>
            public List<string> ImplementsExternal = new List<string>();
            public bool Exposed;
            public bool IsPlainClass => Type == ComponentType.ClassModule && !Exposed;
        }

        private class Unit
        {
            public string Kind;
            public ModuleInfo Module;
            public string Name;
            public List<Declaration> Declarations = new List<Declaration>();
            public HashSet<Unit> Edges = new HashSet<Unit>();
            public bool Live;
            public bool IsModule => Kind == "module";
            public List<string> Roots = new List<string>();
            public List<string> ConditionalRoots = new List<string>();
            public string Display => IsModule ? Module.Name : Module.Name + "." + Name;

            public bool Matches(string name) => Declarations.Any(d => string.Equals(d.IdentifierName, name, StringComparison.OrdinalIgnoreCase));

            /// <summary>Members that something can reach by name at run time: a call by name, a formula, a late-bound access.</summary>
            public bool IsNameable => Kind == "procedure" || Kind == "variable" || Kind == "constant" || Kind == "declare";

            public IEnumerable<string> Names => Declarations.Select(d => d.IdentifierName).Distinct(StringComparer.OrdinalIgnoreCase);
        }

        #endregion

        public ShipAnalysisResult Analyze(IVBProject project, ShipWorkingSet set, ShipTreeShakingSettings settings, out string whyNot)
        {
            whyNot = null;
            settings = settings ?? new ShipTreeShakingSettings();

            ParserState status;
            try
            {
                status = _catchUp();
            }
            catch (ApiException exception)
            {
                whyNot = "Rubberduck could not bring the parser up to date: " + exception.Message;
                return null;
            }

            if (status != ParserState.Ready)
            {
                whyNot = $"The parser is {status}, not Ready, so what is reachable cannot be known." + FailedModules(project.ProjectId);
                return null;
            }

            //The parser state is read below without the editor's thread; if it is re-parsed meanwhile (or the enumeration breaks because of it) the answer is not trusted.
            const string Changed = "the project changed while Ship was working; try again";
            try
            {
                var before = _generation();
                var result = AnalyzeReady(project, set, settings, out whyNot);
                if (result != null && _generation() != before)
                {
                    whyNot = Changed;
                    return null;
                }

                return result;
            }
            catch (InvalidOperationException)
            {
                whyNot = Changed;
                return null;
            }
        }

        private ShipAnalysisResult AnalyzeReady(IVBProject project, ShipWorkingSet set, ShipTreeShakingSettings settings, out string whyNot)
        {
            whyNot = null;
            var projectId = project.ProjectId;
            var failed = FailedModules(projectId);
            if (failed.Length > 0)
            {
                whyNot = "A module of the project did not parse or resolve." + failed;
                return null;
            }

            var all = _state.AllUserDeclarations.Where(d => d.ProjectId == projectId).ToList();
            var moduleDeclarations = new Dictionary<string, Declaration>(StringComparer.OrdinalIgnoreCase);
            foreach (var declaration in all.Where(d => d.DeclarationType.HasFlag(DeclarationType.Module)))
            {
                moduleDeclarations[declaration.IdentifierName] = declaration;
            }

            var infos = new Dictionary<string, ModuleInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (var module in set.Modules.Where(m => !m.Removed))
            {
                if (!moduleDeclarations.TryGetValue(module.Name, out var declaration))
                {
                    whyNot = $"Module {module.Name} is not in the parser state (it is new, or the parser has not read it).";
                    return null;
                }

                var info = new ModuleInfo { Name = module.Name, Type = declaration.QualifiedModuleName.ComponentType, Declaration = declaration };
                info.Exposed = info.Type == ComponentType.ClassModule && declaration is ClassModuleDeclaration classModule && classModule.IsExposed;
                info.Self = new Unit { Kind = "module", Module = info, Name = module.Name };
                info.Self.Declarations.Add(declaration);
                infos[module.Name] = info;
            }

            var unitOfDeclaration = new Dictionary<Declaration, Unit>();
            var ownerOfContext = new Dictionary<ParserRuleContext, Unit>();
            BuildUnits(all, infos, unitOfDeclaration, ownerOfContext);
            AddReferenceEdges(all, infos, projectId, unitOfDeclaration, ownerOfContext);
            AddUnresolvedMemberEdges(infos, projectId, ownerOfContext);
            AddImplements(infos);

            var memberNames = new HashSet<string>(infos.Values.SelectMany(i => i.Units).Where(u => u.IsNameable).SelectMany(u => u.Names), StringComparer.OrdinalIgnoreCase);

            var literals = ReadStringLiterals(set, projectId, infos, out var literalProblem);
            if (literalProblem != null)
            {
                whyNot = literalProblem;
                return null;
            }

            List<PackageRootHit> packageHits;
            try
            {
                packageHits = ReadPackage(set, memberNames);
            }
            catch (Exception exception) when (exception is System.IO.InvalidDataException || exception is System.IO.IOException || exception is UnauthorizedAccessException)
            {
                whyNot = "The document's file could not be read for ribbon callbacks, shape macros and formulas: " + exception.Message;
                return null;
            }

            var result = new ShipAnalysisResult();
            AddRoots(all, infos, settings, literals, packageHits, result.Warnings);
            Propagate(infos);

            ListRoots(infos, literals, packageHits, result);
            ListRemovals(infos, result);
            result.DynamicCalls.AddRange(FindDynamicCalls(all, infos, projectId, ownerOfContext));
            return result;
        }

        #region Parser state checks

        private string FailedModules(string projectId)
        {
            var lines = new List<string>();
            foreach (var pair in _state.ModuleExceptions.Where(e => e.Item1.ProjectId == projectId))
            {
                lines.Add($"{pair.Item1.ComponentName}: {pair.Item2.Message}");
            }

            foreach (var module in _state.DeclarationFinder.AllModules.Where(m => m.ProjectId == projectId))
            {
                var state = _state.GetModuleState(module);
                if (state == ParserState.Error || state == ParserState.ResolverError || state == ParserState.UnexpectedError)
                {
                    lines.Add($"{module.ComponentName}: {state}");
                }
            }

            return lines.Count == 0 ? string.Empty : " " + string.Join("; ", lines.Distinct());
        }

        #endregion

        #region Units and edges

        private static T Up<T>(ParserRuleContext context) where T : ParserRuleContext
        {
            for (var current = context; current != null; current = current.Parent as ParserRuleContext)
            {
                if (current is T found)
                {
                    return found;
                }
            }

            return null;
        }

        private static string ImplementedInterfaceIgnoringCase(ModuleInfo info, string procedureName)
        {
            var separator = procedureName.LastIndexOf('_');
            if (separator <= 0)
            {
                return null;
            }

            var prefix = procedureName.Substring(0, separator);
            var member = procedureName.Substring(separator + 1);
            return info.Implements
                .Where(i => string.Equals(i.Name, prefix, StringComparison.OrdinalIgnoreCase) && i.Units.Any(u => u.Kind == "procedure" && u.Matches(member)))
                .Select(i => i.Name)
                .Concat(info.ImplementsExternal.Where(n => string.Equals(n, prefix, StringComparison.OrdinalIgnoreCase)))
                .FirstOrDefault();
        }

        private static bool IsModuleLevelParent(Declaration declaration)
            => declaration.ParentDeclaration != null && declaration.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Module);

        private void BuildUnits(List<Declaration> all, Dictionary<string, ModuleInfo> infos, Dictionary<Declaration, Unit> unitOfDeclaration, Dictionary<ParserRuleContext, Unit> ownerOfContext)
        {
            var byKey = new Dictionary<string, Unit>(StringComparer.OrdinalIgnoreCase);
            var byStatement = new Dictionary<ParserRuleContext, Unit>();

            Unit Get(ModuleInfo info, string kind, string name, string key)
            {
                var full = info.Name + "|" + key;
                if (!byKey.TryGetValue(full, out var unit))
                {
                    unit = new Unit { Kind = kind, Module = info, Name = name };
                    byKey[full] = unit;
                    info.Units.Add(unit);
                }

                return unit;
            }

            foreach (var declaration in all.OrderBy(d => d.QualifiedModuleName.ComponentName).ThenBy(d => d.Selection.StartLine).ThenBy(d => d.Selection.StartColumn))
            {
                if (!infos.TryGetValue(declaration.QualifiedModuleName.ComponentName, out var info))
                {
                    continue;
                }

                var type = declaration.DeclarationType;
                if (type.HasFlag(DeclarationType.Module))
                {
                    unitOfDeclaration[declaration] = info.Self;
                    continue;
                }

                if (!IsModuleLevelParent(declaration))
                {
                    continue;
                }

                Unit unit = null;
                if (type == DeclarationType.LibraryFunction || type == DeclarationType.LibraryProcedure)
                {
                    unit = Get(info, "declare", declaration.IdentifierName, "d:" + declaration.IdentifierName);
                }
                else if (type.HasFlag(DeclarationType.Member))
                {
                    unit = Get(info, "procedure", declaration.IdentifierName, "m:" + declaration.IdentifierName);
                }
                else if (type == DeclarationType.Variable || type == DeclarationType.Constant)
                {
                    var statement = type == DeclarationType.Variable
                        ? (ParserRuleContext)Up<VBAParser.VariableStmtContext>(declaration.Context)
                        : Up<VBAParser.ConstStmtContext>(declaration.Context);
                    if (statement != null && byStatement.TryGetValue(statement, out var existing))
                    {
                        unit = existing;
                    }
                    else
                    {
                        unit = Get(info, type == DeclarationType.Variable ? "variable" : "constant", declaration.IdentifierName, "v:" + declaration.IdentifierName);
                        if (statement != null)
                        {
                            byStatement[statement] = unit;
                        }
                    }
                }
                else if (type == DeclarationType.UserDefinedType)
                {
                    unit = Get(info, "type", declaration.IdentifierName, "t:" + declaration.IdentifierName);
                }
                else if (type == DeclarationType.Enumeration)
                {
                    unit = Get(info, "enum", declaration.IdentifierName, "e:" + declaration.IdentifierName);
                }

                if (unit == null)
                {
                    continue;
                }

                unit.Declarations.Add(declaration);
                unitOfDeclaration[declaration] = unit;
                for (var context = declaration.Context; context != null && !(context is VBAParser.ModuleDeclarationsContext || context is VBAParser.ModuleBodyContext || context is VBAParser.ModuleContext);
                     context = context.Parent as ParserRuleContext)
                {
                    ownerOfContext[context] = unit;
                }
            }
        }

        private static Unit OwnerOf(IdentifierReference reference, ModuleInfo module, Dictionary<ParserRuleContext, Unit> ownerOfContext)
            => OwnerOf(reference.Context, module, ownerOfContext);

        private static Unit OwnerOf(ParserRuleContext start, ModuleInfo module, Dictionary<ParserRuleContext, Unit> ownerOfContext)
        {
            for (var context = start; context != null; context = context.Parent as ParserRuleContext)
            {
                if (ownerOfContext.TryGetValue(context, out var unit))
                {
                    return unit;
                }
            }

            return module.Self;
        }

        private static Unit TargetOf(Declaration declaration, Dictionary<Declaration, Unit> unitOfDeclaration)
        {
            for (var current = declaration; current != null; current = current.ParentDeclaration)
            {
                if (unitOfDeclaration.TryGetValue(current, out var unit))
                {
                    return unit;
                }
            }

            return null;
        }

        private void AddReferenceEdges(List<Declaration> all, Dictionary<string, ModuleInfo> infos, string projectId, Dictionary<Declaration, Unit> unitOfDeclaration, Dictionary<ParserRuleContext, Unit> ownerOfContext)
        {
            foreach (var declaration in all)
            {
                var target = TargetOf(declaration, unitOfDeclaration);
                if (target == null)
                {
                    continue;
                }

                foreach (var reference in declaration.References)
                {
                    if (reference.QualifiedModuleName.ProjectId != projectId || !infos.TryGetValue(reference.QualifiedModuleName.ComponentName, out var module))
                    {
                        continue;
                    }

                    var owner = OwnerOf(reference, module, ownerOfContext);
                    if (owner != target)
                    {
                        owner.Edges.Add(target);
                    }
                }
            }
        }

        //A call on an object the parser could not type (Object, Variant) is late bound: any member of that name may be the one that runs.
        private void AddUnresolvedMemberEdges(Dictionary<string, ModuleInfo> infos, string projectId, Dictionary<ParserRuleContext, Unit> ownerOfContext)
        {
            var byName = infos.Values.SelectMany(i => i.Units).Where(u => u.Kind == "procedure" || u.Kind == "variable" || u.Kind == "constant")
                .SelectMany(u => u.Names.Select(n => (name: n, unit: u)))
                .GroupBy(x => x.name, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.Select(x => x.unit).Distinct().ToList(), StringComparer.OrdinalIgnoreCase);

            foreach (var unresolved in _state.DeclarationFinder.UnresolvedMemberDeclarations().Where(u => u.ProjectId == projectId))
            {
                if (!byName.TryGetValue(unresolved.IdentifierName, out var targets))
                {
                    continue;
                }

                if (!infos.TryGetValue(unresolved.QualifiedModuleName.ComponentName, out var module))
                {
                    continue;
                }

                //The unbound member is a declaration of the place it is used in; it has no references of its own.
                var owner = OwnerOf(unresolved.Context, module, ownerOfContext);
                foreach (var target in targets.Where(t => t != owner))
                {
                    owner.Edges.Add(target);
                }
            }
        }

        private static void AddImplements(Dictionary<string, ModuleInfo> infos)
        {
            foreach (var info in infos.Values.Where(i => i.Type == ComponentType.ClassModule))
            {
                if (info.Declaration is ClassModuleDeclaration classModule)
                {
                    info.ImplementsExternal.AddRange(classModule.Supertypes.Where(s => !s.IsUserDefined).Select(s => s.IdentifierName));
                }

                foreach (var reference in info.Declaration.References)
                {
                    if (Up<VBAParser.ImplementsStmtContext>(reference.Context) != null
                        && infos.TryGetValue(reference.QualifiedModuleName.ComponentName, out var implementor)
                        && !implementor.Implements.Contains(info))
                    {
                        implementor.Implements.Add(info);
                    }
                }
            }
        }

        #endregion

        #region Roots

        private class Literals
        {
            /// <summary>Name -> the modules with a string literal that holds it as a whole word.</summary>
            public Dictionary<string, HashSet<string>> Modules = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            /// <summary>Name -> the modules that have it as an identifier inside an #If, #ElseIf or #Else branch (active or not).</summary>
            public Dictionary<string, HashSet<string>> Conditional = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            public static void Add(Dictionary<string, HashSet<string>> target, string name, string module)
            {
                if (!target.TryGetValue(name, out var modules))
                {
                    target[name] = modules = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                }

                modules.Add(module);
            }
        }

        //Words of a string literal: text of formulas and macro names, not VBA, so it is split as text.
        private static readonly Regex Word = new Regex(@"[\p{L}_][\p{L}\p{N}_]*", RegexOptions.Compiled);

        /// <summary>
        /// Every identifier-like token of the lines between a conditional compilation #If and its #End If, active branch or not, nested blocks included. The parser of the code only sees
        /// the active branch, so a reference in another branch makes no edge: these names are kept instead. Found with Rubberduck's own lexer and its conditional compilation grammar
        /// (the one its preprocessor uses), so a #If inside a string or a comment is not a block. Comments inside a block are not counted; the words of a string literal are.
        /// </summary>
        internal static IEnumerable<string> ConditionalWords(string text)
        {
            var tokens = new SimpleVBAModuleTokenStreamProvider().Tokens(text ?? string.Empty);
            var parser = new VBAConditionalCompilationParser(tokens);
            parser.RemoveErrorListeners();
            var finder = new IfBlockFinder();
            ParseTreeWalker.Default.Walk(finder, parser.compilationUnit());

            var all = tokens.GetTokens();
            foreach (var (from, to) in finder.Blocks)
            {
                var inComment = false;
                for (var i = from; i <= to && i < all.Count; i++)
                {
                    var token = all[i];
                    switch (token.Type)
                    {
                        case VBALexer.NEWLINE:
                            inComment = false;
                            continue;
                        case VBALexer.SINGLEQUOTE:
                        case VBALexer.REM:
                            inComment = true;
                            continue;
                    }

                    if (inComment)
                    {
                        continue;
                    }

                    if (token.Type == VBALexer.STRINGLITERAL)
                    {
                        foreach (var candidate in Candidates(Unquote(token.Text)))
                        {
                            yield return candidate;
                        }
                    }
                    else if (!string.IsNullOrEmpty(token.Text) && (char.IsLetter(token.Text[0]) || token.Text[0] == '_'))
                    {
                        yield return token.Text;
                    }
                }
            }
        }

        private class IfBlockFinder : VBAConditionalCompilationParserBaseListener
        {
            private int _depth;
            public List<(int From, int To)> Blocks { get; } = new List<(int, int)>();

            public override void EnterCcIfBlock(VBAConditionalCompilationParser.CcIfBlockContext context)
            {
                if (_depth++ == 0 && context.Start != null && context.Stop != null)
                {
                    Blocks.Add((context.Start.TokenIndex, context.Stop.TokenIndex));
                }
            }

            public override void ExitCcIfBlock(VBAConditionalCompilationParser.CcIfBlockContext context) => _depth--;
        }

        private static string Unquote(string literal)
            => literal.Length >= 2 && literal[0] == '"' && literal[literal.Length - 1] == '"' ? literal.Substring(1, literal.Length - 2).Replace("\"\"", "\"") : literal;

        private Literals ReadStringLiterals(ShipWorkingSet set, string projectId, Dictionary<string, ModuleInfo> infos, out string problem)
        {
            problem = null;
            var literals = new Literals();
            foreach (var module in set.Modules.Where(m => !m.Removed && infos.ContainsKey(m.Name)))
            {
                IParseTree tree;
                try
                {
                    (tree, _) = _parser.Parse(module.Name, projectId, module.Text ?? string.Empty, CancellationToken.None, CodeKind.AttributesCode);
                }
                catch (Exception exception) when (!(exception is OperationCanceledException))
                {
                    problem = $"Module {module.Name} could not be read for string literals: {exception.Message}";
                    return null;
                }

                var finder = new LiteralFinder();
                ParseTreeWalker.Default.Walk(finder, tree);
                foreach (var text in finder.Texts)
                {
                    foreach (var candidate in Candidates(text))
                    {
                        Literals.Add(literals.Modules, candidate, module.Name);
                    }
                }

                foreach (var word in ConditionalWords(module.Text))
                {
                    Literals.Add(literals.Conditional, word, module.Name);
                }
            }

            return literals;
        }

        //"Foo", "Module1.Foo", "'Book.xlam'!Module1.Foo": the whole text and the part after the last '!' or '.'; and every word in it, for "=MyUdf(A1)", Evaluate("MyUdf(2)") and "'Mod.Foo 1'".
        private static IEnumerable<string> Candidates(string text)
        {
            text = (text ?? string.Empty).Trim();
            if (text.Length == 0)
            {
                yield break;
            }

            yield return text;
            foreach (Match match in Word.Matches(text))
            {
                yield return match.Value;
            }

            var cut = Math.Max(text.LastIndexOf('!'), text.LastIndexOf('.'));
            if (cut >= 0 && cut < text.Length - 1)
            {
                yield return text.Substring(cut + 1).Trim();
            }
        }

        private class LiteralFinder : VBAParserBaseListener
        {
            public List<string> Texts { get; } = new List<string>();

            public override void EnterLiteralExpression(VBAParser.LiteralExpressionContext context)
            {
                var literal = context.STRINGLITERAL();
                if (literal == null || Up<VBAParser.AttributeStmtContext>(context) != null)
                {
                    return;
                }

                var text = literal.GetText();
                if (text.Length >= 2 && text[0] == '"' && text[text.Length - 1] == '"')
                {
                    text = text.Substring(1, text.Length - 2).Replace("\"\"", "\"");
                }

                Texts.Add(text);
            }
        }

        private static List<PackageRootHit> ReadPackage(ShipWorkingSet set, IEnumerable<string> memberNames)
        {
            if (set.OpenPackage == null)
            {
                return new List<PackageRootHit>();
            }

            using (var stream = set.OpenPackage())
            {
                return stream == null ? new List<PackageRootHit>() : PackageRoots.Find(stream, memberNames);
            }
        }

        private static bool IsPublic(Declaration declaration)
            => declaration.Accessibility == Accessibility.Public || declaration.Accessibility == Accessibility.Global || declaration.Accessibility == Accessibility.Implicit;

        private void AddRoots(List<Declaration> all, Dictionary<string, ModuleInfo> infos, ShipTreeShakingSettings settings, Literals literals, List<PackageRootHit> packageHits, List<string> warnings)
        {
            var finder = _state.DeclarationFinder;
            var handlers = new HashSet<Declaration>(finder.FindEventHandlers());
            var conservative = settings.Mode == ShipTreeShakingMode.Conservative;

            foreach (var info in infos.Values)
            {
                foreach (var unit in info.Units)
                {
                    if (info.Type == ComponentType.UserForm)
                    {
                        unit.Roots.Add("UserForm code is never removed");
                        continue;
                    }

                    //The default member (coll(1), x = coll) and the enumerator (For Each) are called without their names.
                    if (unit.Declarations.Any(d => d.Attributes.Any(a => a.Name.Equals("VB_UserMemId", StringComparison.OrdinalIgnoreCase)
                                                                         || a.Name.EndsWith(".VB_UserMemId", StringComparison.OrdinalIgnoreCase))))
                    {
                        unit.ConditionalRoots.Add("default or enumerator member (VB_UserMemId)");
                    }

                    if (unit.Kind != "procedure")
                    {
                        continue;
                    }

                    foreach (var declaration in unit.Declarations)
                    {
                        var conditional = info.IsPlainClass;
                        if (declaration.Annotations.Any(a => a.Annotation is EntryPointAnnotation || a.Annotation is ExcelHotKeyAnnotation))
                        {
                            unit.Roots.Add("@EntryPoint / @ExcelHotkey");
                        }

                        if (handlers.Contains(declaration) || (info.Type != ComponentType.StandardModule && DocumentHandlerPrefixes.Any(p => declaration.IdentifierName.StartsWith(p, StringComparison.OrdinalIgnoreCase))))
                        {
                            (conditional ? unit.ConditionalRoots : unit.Roots).Add("event handler");
                        }

                        //Controls on a worksheet are not declared where the parser can see them, and their handlers (Button1_Click) are not found by name: a Name_Event procedure of a document module stays.
                        if (info.Type == ComponentType.Document && declaration.IdentifierName.Contains("_") && !handlers.Contains(declaration))
                        {
                            unit.Roots.Add("possible event handler of a control on the document");
                        }

                        if (declaration is ModuleBodyElementDeclaration element)
                        {
                            //The parser matches IUnitOfWork_Command to the interface member command case-sensitively, the VBE does not: a different case is still an implementation.
                            var implemented = element.IsInterfaceImplementation
                                ? element.InterfaceImplemented?.IdentifierName
                                : ImplementedInterfaceIgnoringCase(info, element.IdentifierName);
                            if (implemented != null)
                            {
                                (conditional ? unit.ConditionalRoots : unit.Roots).Add("interface implementation of " + implemented);
                            }
                        }

                        if (info.Type == ComponentType.StandardModule)
                        {
                            if (AutoMacros.Any(a => a.Equals(declaration.IdentifierName, StringComparison.OrdinalIgnoreCase)))
                            {
                                unit.Roots.Add("Auto macro run by Excel");
                            }

                            if (conservative && IsPublic(declaration))
                            {
                                unit.Roots.Add("public procedure of a standard module");
                            }
                        }
                        else if (info.Exposed && IsPublic(declaration))
                        {
                            unit.Roots.Add("public member of a class with VB_Exposed = True");
                        }
                    }
                }
            }

            foreach (var field in finder.FindWithEventFields())
            {
                if (field.ProjectId == null || !infos.TryGetValue(field.QualifiedModuleName.ComponentName, out var info)
                    || !finder.FindHandlersForWithEventsField(field).Any())
                {
                    continue;
                }

                var unit = info.Units.FirstOrDefault(u => u.Declarations.Contains(field));
                unit?.ConditionalRoots.Add("WithEvents variable that has handlers");
            }

            foreach (var entry in settings.Keep.Where(k => !string.IsNullOrWhiteSpace(k)).Select(k => k.Trim()))
            {
                var dot = entry.IndexOf('.');
                var moduleName = dot > 0 ? entry.Substring(0, dot) : null;
                var name = dot > 0 ? entry.Substring(dot + 1) : entry;
                var matched = false;
                foreach (var info in infos.Values)
                {
                    //"Logger" and "Logger.Logger" name the module itself.
                    if (string.Equals(info.Name, name, StringComparison.OrdinalIgnoreCase) && (moduleName == null || string.Equals(info.Name, moduleName, StringComparison.OrdinalIgnoreCase)))
                    {
                        matched = true;
                        info.Self.Roots.Add("keep list (" + entry + ")");
                    }

                    foreach (var unit in info.Units)
                    {
                        var matches = moduleName == null
                            ? unit.Matches(name) || string.Equals(info.Name, name, StringComparison.OrdinalIgnoreCase)
                            : string.Equals(info.Name, moduleName, StringComparison.OrdinalIgnoreCase) && unit.Matches(name);
                        if (matches)
                        {
                            matched = true;
                            unit.Roots.Add("keep list (" + entry + ")");
                        }
                    }
                }

                if (!matched)
                {
                    warnings.Add($"Tree shaking: keep entry {entry} matches nothing.");
                }
            }

            //A module is named in a string or in a conditional branch (Dim x As Logger in a branch the parser did not read): it is not unreachable.
            foreach (var info in infos.Values)
            {
                var inStrings = ModulesNaming(info.Name, literals.Modules, info.Name);
                if (inStrings.Count > 0)
                {
                    info.Self.Roots.Add("string literal in " + string.Join(", ", inStrings));
                }

                var inBranches = ModulesNaming(info.Name, literals.Conditional, info.Name);
                if (inBranches.Count > 0)
                {
                    info.Self.Roots.Add("name inside a conditional compilation block in " + string.Join(", ", inBranches));
                }
            }

            foreach (var unit in infos.Values.SelectMany(i => i.Units))
            {
                var inLiterals = LiteralModules(unit, literals.Modules);
                if (unit.IsNameable && inLiterals.Count > 0)
                {
                    unit.Roots.Add("string literal in " + string.Join(", ", inLiterals));
                }

                var inBranches = LiteralModules(unit, literals.Conditional);
                if (inBranches.Count > 0)
                {
                    unit.Roots.Add("name inside a conditional compilation block in " + string.Join(", ", inBranches));
                }

                var hit = unit.IsNameable ? packageHits.FirstOrDefault(h => unit.Names.Any(n => string.Equals(h.Name, n, StringComparison.OrdinalIgnoreCase))) : null;
                if (hit != null)
                {
                    unit.Roots.Add("found in " + hit.Source);
                }
            }
        }

        private static List<string> ModulesNaming(string name, Dictionary<string, HashSet<string>> source, string except)
            => source.TryGetValue(name, out var modules)
                ? modules.Where(m => !string.Equals(m, except, StringComparison.OrdinalIgnoreCase)).OrderBy(m => m, StringComparer.OrdinalIgnoreCase).ToList()
                : new List<string>();

        private static List<string> LiteralModules(Unit unit, Dictionary<string, HashSet<string>> source)
            => unit.Names.Where(source.ContainsKey).SelectMany(n => source[n]).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(m => m, StringComparer.OrdinalIgnoreCase).ToList();

        private static void ListRoots(Dictionary<string, ModuleInfo> infos, Literals literals, List<PackageRootHit> packageHits, ShipAnalysisResult result)
        {
            var lines = new List<string>();
            foreach (var info in infos.Values.OrderBy(i => i.Name, StringComparer.OrdinalIgnoreCase))
            {
                if (info.Type == ComponentType.UserForm)
                {
                    lines.Add($"{info.Name}: UserForm, its code is never removed");
                    continue;
                }

                if (info.Exposed)
                {
                    lines.Add($"{info.Name}: class with VB_Exposed = True");
                }

                foreach (var reason in info.Self.Roots.Distinct())
                {
                    lines.Add($"{info.Name}: {reason}");
                }

                foreach (var unit in info.Units)
                {
                    foreach (var reason in unit.Roots.Distinct())
                    {
                        lines.Add($"{unit.Display}: {reason}");
                    }

                    foreach (var reason in unit.ConditionalRoots.Distinct())
                    {
                        lines.Add($"{unit.Display}: {reason} (kept while {info.Name} is)");
                    }
                }
            }

            result.Roots.AddRange(lines.Distinct().OrderBy(l => l, StringComparer.OrdinalIgnoreCase));
        }

        #endregion

        #region Reachability

        private static void Propagate(Dictionary<string, ModuleInfo> infos)
        {
            var queue = new Queue<Unit>();

            void Mark(Unit unit)
            {
                if (!unit.Live)
                {
                    unit.Live = true;
                    queue.Enqueue(unit);
                }
            }

            foreach (var info in infos.Values)
            {
                if (info.Exposed || info.Type == ComponentType.Document || info.Type == ComponentType.UserForm)
                {
                    Mark(info.Self);
                }

                if (info.Self.Roots.Count > 0)
                {
                    Mark(info.Self);
                }

                foreach (var unit in info.Units.Where(u => u.Roots.Count > 0))
                {
                    Mark(unit);
                }
            }

            while (queue.Count > 0)
            {
                var unit = queue.Dequeue();
                if (unit.IsModule)
                {
                    var info = unit.Module;
                    //A class that implements an interface has to implement all of it, so all of the interface stays.
                    foreach (var implemented in info.Implements)
                    {
                        Mark(implemented.Self);
                        foreach (var member in implemented.Units.Where(u => u.Kind == "procedure"))
                        {
                            Mark(member);
                        }
                    }

                    foreach (var conditional in info.Units.Where(u => u.ConditionalRoots.Count > 0))
                    {
                        Mark(conditional);
                    }
                }
                else
                {
                    Mark(unit.Module.Self);
                }

                foreach (var target in unit.Edges)
                {
                    Mark(target);
                }
            }
        }

        private static void ListRemovals(Dictionary<string, ModuleInfo> infos, ShipAnalysisResult result)
        {
            foreach (var info in infos.Values.OrderBy(i => i.Name, StringComparer.OrdinalIgnoreCase))
            {
                if (info.Type == ComponentType.UserForm)
                {
                    continue;
                }

                foreach (var unit in info.Units.Where(u => !u.Live))
                {
                    foreach (var name in unit.Declarations.Select(d => d.IdentifierName).Distinct(StringComparer.OrdinalIgnoreCase).Where(n => unit.Kind == "variable" || unit.Kind == "constant" || n == unit.Name))
                    {
                        result.Removals.Add(new ShipRemoval { Module = info.Name, Name = name, Kind = unit.Kind });
                    }
                }

                if (!info.Self.Live && (info.Type == ComponentType.StandardModule || info.Type == ComponentType.ClassModule))
                {
                    result.Removals.Add(new ShipRemoval { Module = info.Name, Name = info.Name, Kind = "module" });
                }
            }
        }

        #endregion

        #region Dynamic calls

        private IEnumerable<string> FindDynamicCalls(List<Declaration> all, Dictionary<string, ModuleInfo> infos, string projectId, Dictionary<ParserRuleContext, Unit> ownerOfContext)
        {
            var lines = new List<string>();
            foreach (var callee in _state.DeclarationFinder.AllBuiltInDeclarations.Where(IsDynamicCallee))
            {
                foreach (var reference in callee.References)
                {
                    if (reference.QualifiedModuleName.ProjectId != projectId || !infos.TryGetValue(reference.QualifiedModuleName.ComponentName, out var module))
                    {
                        continue;
                    }

                    var owner = OwnerOf(reference, module, ownerOfContext);
                    if (!owner.Live && !owner.IsModule)
                    {
                        continue;
                    }

                    var index = callee.IdentifierName.Equals("Run", StringComparison.OrdinalIgnoreCase) ? 0 : 1;
                    var argument = ArgumentAt(reference, index);
                    if (argument == null || IsStringLiteral(argument))
                    {
                        continue;
                    }

                    var where = owner.IsModule ? module.Name + ".(declarations)" : owner.Display;
                    var call = callee.IdentifierName.Equals("CallByName", StringComparison.OrdinalIgnoreCase) ? "CallByName" : "Application." + callee.IdentifierName;
                    lines.Add($"{where}: {call} with a target that is not a string literal ({argument.GetText().Trim()}), line {reference.Selection.StartLine}");
                }
            }

            return lines.Distinct().OrderBy(l => l, StringComparer.OrdinalIgnoreCase);
        }

        private static bool IsDynamicCallee(Declaration declaration)
        {
            if (declaration.IsUserDefined)
            {
                return false;
            }

            var name = declaration.IdentifierName;
            if (name.Equals("CallByName", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            var parent = declaration.ParentDeclaration?.IdentifierName;
            return (name.Equals("Run", StringComparison.OrdinalIgnoreCase) || name.Equals("OnTime", StringComparison.OrdinalIgnoreCase) || name.Equals("OnKey", StringComparison.OrdinalIgnoreCase))
                   && (string.Equals(parent, "Application", StringComparison.OrdinalIgnoreCase) || string.Equals(parent, "_Application", StringComparison.OrdinalIgnoreCase));
        }

        //The expression of the argument at the index (named arguments count by position), or null if the call has none there.
        private static VBAParser.ExpressionContext ArgumentAt(IdentifierReference reference, int index)
        {
            var expression = Up<VBAParser.LExpressionContext>(reference.Context);
            VBAParser.ArgumentListContext arguments = null;
            switch (expression?.Parent)
            {
                case VBAParser.IndexExprContext indexed when indexed.lExpression() == expression:
                    arguments = indexed.argumentList();
                    break;
                case VBAParser.CallStmtContext call when call.lExpression() == expression:
                    arguments = call.argumentList();
                    break;
            }

            var list = arguments?.argument();
            if (list == null || list.Length <= index)
            {
                return null;
            }

            var argument = list[index];
            return argument.positionalArgument()?.argumentExpression()?.expression()
                   ?? argument.namedArgument()?.argumentExpression()?.expression();
        }

        private static bool IsStringLiteral(VBAParser.ExpressionContext expression)
            => expression is VBAParser.LiteralExprContext literal && literal.literalExpression()?.STRINGLITERAL() != null;

        #endregion
    }
}
