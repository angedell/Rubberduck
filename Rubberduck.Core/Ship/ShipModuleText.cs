using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Parsing.VBA.Parsing.ParsingExceptions;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Ship
{
    /// <summary>The outcome of an edit of a module's text. A module that does not parse comes back unchanged with <see cref="Parsed"/> false.</summary>
    public class ShipTextResult
    {
        public string Text { get; set; }
        public bool Parsed { get; set; } = true;
        /// <summary>Why the module did not parse: the parser's complaint with its place.</summary>
        public string Problem { get; set; }
        /// <summary>How many comments or attribute lines were taken out.</summary>
        public int Removed { get; set; }
        /// <summary>How many lines got a line number.</summary>
        public int Numbered { get; set; }
        /// <summary>Procedures left alone because they have line numbers of their own.</summary>
        public List<string> Skipped { get; } = new List<string>();
    }

    /// <summary>
    /// Edits the text of an exported module (header, attributes and code) with the parser: the token stream of a parse is rewritten by deleting the parse tree's own
    /// comment and attribute contexts, so what a string literal, a <c>Rem</c>, a line continuation or a conditional compilation line means is the parser's call, never a pattern's.
    /// </summary>
    /// <remarks>
    /// What the parser does, and so what this does (each is covered by a test):
    /// <list type="bullet">
    /// <item>A comment is part of a logical line, as in the language specification: a comment that ends in a line continuation (<c> _</c>) goes on into the next physical line.</item>
    /// <item>The preprocessor puts the tokens of an inactive <c>#If</c> branch, and of every <c>#If</c>, <c>#ElseIf</c>, <c>#Else</c> and <c>#End If</c> line itself, on the hidden channel.
    /// They are not part of the parse tree, so there is no comment context for them; they are found as comment-mark tokens instead.</item>
    /// <item>The grammar takes <c>Rem</c> after a colon in a one-line <c>If ... Then</c> for a call of a procedure called Rem, so the tree has no comment there either; the same token rule finds it.</item>
    /// <item>The header of a class or form (<c>VERSION</c>, <c>BEGIN ... END</c>) is the designer's, and its comments (<c>MultiUse = -1  'True</c>) are left alone.</item>
    /// </list>
    /// </remarks>
    public class ShipModuleText
    {
        private readonly IStringParser _parser;

        public ShipModuleText(IStringParser parser)
        {
            _parser = parser;
        }

        /// <summary>Removes every comment and annotation, and the lines that held nothing else.</summary>
        public ShipTextResult StripComments(string moduleName, string projectId, string text)
        {
            return Edit(moduleName, projectId, text, (tree, tokens) =>
            {
                var finder = new CommentFinder();
                ParseTreeWalker.Default.Walk(finder, tree);
                var spans = finder.Spans.ToList();
                spans.AddRange(CommentMarks(tokens, finder.Headers));
                return spans;
            });
        }

        /// <summary>
        /// Removes the attribute lines the shipped project does not need. Kept: <c>VB_Name</c>, <c>VB_PredeclaredId</c>, <c>VB_Exposed</c>, <c>VB_Creatable</c>, <c>VB_GlobalNameSpace</c>,
        /// <c>VB_UserMemId</c>, and <c>VB_Description</c> and <c>VB_ProcData.VB_Invoke_Func</c> of public procedures in a standard module without <c>Option Private Module</c>
        /// (Excel shows those in the macro and function lists). Everything else, every <c>VB_Ext_KEY</c> included, goes.
        /// </summary>
        public ShipTextResult StripDescriptiveAttributes(string moduleName, string projectId, string text, ComponentType type)
        {
            return Edit(moduleName, projectId, text, (tree, tokens) =>
            {
                var finder = new AttributeFinder(type);
                ParseTreeWalker.Default.Walk(finder, tree);
                return finder.Spans;
            });
        }

        /// <summary>The module's code alone: no header, no designer block and no attribute lines, as the code pane of the VBE holds it.</summary>
        public ShipTextResult CodeOnly(string moduleName, string projectId, string text)
        {
            return Edit(moduleName, projectId, text, (tree, tokens) =>
            {
                var finder = new HeaderFinder();
                ParseTreeWalker.Default.Walk(finder, tree);
                return finder.Spans;
            });
        }

        private ShipTextResult Edit(string moduleName, string projectId, string text, Func<IParseTree, IList<IToken>, List<Span>> find)
        {
            text = text ?? string.Empty;
            if (!TryParse(moduleName, projectId, text, out var tree, out var stream, out var problem))
            {
                return Unparsed(text, problem);
            }

            var tokens = ((BufferedTokenStream)stream).GetTokens();
            var spans = find(tree, tokens).Where(s => s.To >= s.From).OrderBy(s => s.From).ToList();
            var merged = new List<Span>();
            foreach (var span in spans)
            {
                if (merged.Count > 0 && span.From <= merged[merged.Count - 1].To)
                {
                    var last = merged[merged.Count - 1];
                    merged[merged.Count - 1] = new Span(last.From, Math.Max(last.To, span.To));
                }
                else
                {
                    merged.Add(span);
                }
            }

            if (merged.Count == 0)
            {
                return new ShipTextResult { Text = text };
            }

            var rewriter = new TokenStreamRewriter(stream);
            foreach (var span in merged)
            {
                var (from, to) = WholeLines(tokens, span);
                rewriter.Delete(from, to);
            }

            return new ShipTextResult { Text = rewriter.GetText(), Removed = merged.Count };
        }

        /// <summary>
        /// Puts a line number in front of every executable statement that starts a line inside a procedure, so that <c>Erl</c> can say where an error happened.
        /// The number is the line the statement has in the code pane of the module (the header, the designer block and the attribute lines are not counted), which is the line of the
        /// user's own project: it stays true after the steps that follow have taken lines out. Not numbered: declarations (<c>Dim</c>, <c>Static</c>, <c>Const</c>), the lines of
        /// <c>Sub</c>, <c>Function</c>, <c>Property</c>, <c>End</c>, <c>Else</c>, <c>ElseIf</c> and <c>Case</c>, comments, blank lines, continuation lines, and statements that follow a colon.
        /// A procedure that has a line number of its own is left alone: mixing the two would give <c>Erl</c> two meanings.
        /// </summary>
        public ShipTextResult NumberLines(string moduleName, string projectId, string text)
        {
            text = text ?? string.Empty;
            if (!TryParse(moduleName, projectId, text, out var tree, out var stream, out var problem))
            {
                return Unparsed(text, problem);
            }

            var tokens = ((BufferedTokenStream)stream).GetTokens();
            var finder = new LineNumberFinder(tokens);
            ParseTreeWalker.Default.Walk(finder, tree);

            var result = new ShipTextResult { Text = text };
            result.Skipped.AddRange(finder.Skipped);
            var rewriter = new TokenStreamRewriter(stream);
            foreach (var statement in finder.Statements)
            {
                var number = statement.Start.Line - finder.OutsideCodePaneBefore(statement.Start.Line);
                if (number < 1 || number > MaxLineNumber)
                {
                    continue;
                }

                var index = statement.Start.TokenIndex;
                while (index > 0 && tokens[index - 1].Type == VBALexer.WS)
                {
                    index--;
                }

                rewriter.InsertBefore(index, number + " ");
                result.Numbered++;
            }

            if (result.Numbered > 0)
            {
                result.Text = rewriter.GetText();
            }

            return result;
        }

        //VBA keeps line numbers in 0..65535.
        private const int MaxLineNumber = 65535;

        private bool TryParse(string moduleName, string projectId, string text, out IParseTree tree, out ITokenStream stream, out string problem)
        {
            tree = null;
            stream = null;
            problem = null;
            try
            {
                (tree, stream) = _parser.Parse(moduleName, projectId, text, CancellationToken.None, CodeKind.AttributesCode);
                return true;
            }
            catch (SyntaxErrorException exception)
            {
                problem = $"line {exception.LineNumber}, column {exception.Position}: {Short(exception.Message)} (at '{exception.OffendingSymbol?.Text}')";
            }
            catch (Exception exception) when (!(exception is OperationCanceledException))
            {
                problem = exception.Message;
            }

            return false;
        }

        private static ShipTextResult Unparsed(string text, string problem)
            => new ShipTextResult { Text = text, Parsed = false, Problem = problem };

        private static string Short(string message)
        {
            var firstLine = (message ?? string.Empty).Split('\r', '\n')[0];
            return firstLine.Length > 120 ? firstLine.Substring(0, 120) : firstLine;
        }

        //Takes the spaces in front of the span with it, and the line break behind it if that leaves the line empty.
        private static (int From, int To) WholeLines(IList<IToken> tokens, Span span)
        {
            var from = span.From;
            while (from > 0 && tokens[from - 1].Type == VBALexer.WS)
            {
                from--;
            }

            var startsLine = from == 0 || tokens[from - 1].Type == VBALexer.NEWLINE;
            var to = span.To;
            var next = to + 1 < tokens.Count ? tokens[to + 1] : null;
            if (startsLine && next != null && next.Type == VBALexer.NEWLINE)
            {
                to++;
            }

            return (from, to);
        }

        //The parse tree has no comment context for what the preprocessor hid (inactive code and every #If line),
        //and the grammar takes "Rem" after a colon in a one-line If for a call. So every comment mark that is a token of its own is a comment, outside the designer's header:
        //a mark inside a string is part of the string's token, and the lexer has no other use for it.
        private static IEnumerable<Span> CommentMarks(IList<IToken> tokens, IReadOnlyList<Span> headers)
        {
            for (var i = 0; i < tokens.Count; i++)
            {
                var token = tokens[i];
                if (headers.Any(h => i >= h.From && i <= h.To))
                {
                    continue;
                }

                if (token.Type == VBALexer.SINGLEQUOTE || (token.Type == VBALexer.REM && StartsStatement(tokens, i)))
                {
                    var end = i;
                    while (end + 1 < tokens.Count && tokens[end + 1].Type != VBALexer.NEWLINE && tokens[end + 1].Type != TokenConstants.Eof)
                    {
                        end++;
                    }

                    yield return new Span(i, end);
                    i = end;
                }
            }
        }

        private static bool StartsStatement(IList<IToken> tokens, int index)
        {
            for (var i = index - 1; i >= 0; i--)
            {
                var type = tokens[i].Type;
                if (type == VBALexer.WS || type == VBALexer.LINE_CONTINUATION)
                {
                    continue;
                }

                return type == VBALexer.NEWLINE || type == VBALexer.COLON || type == VBALexer.THEN || type == VBALexer.ELSE;
            }

            return true;
        }

        private struct Span
        {
            public Span(int from, int to)
            {
                From = from;
                To = to;
            }

            public int From { get; }
            public int To { get; }
        }

        private static Span Of(ParserRuleContext context)
            => context?.Start == null || context.Stop == null ? new Span(0, -1) : new Span(context.Start.TokenIndex, context.Stop.TokenIndex);

        //The header of a class or form is the designer's; its comments are not ours to remove.
        private class CommentFinder : VBAParserBaseListener
        {
            private int _inHeader;
            public List<Span> Spans { get; } = new List<Span>();
            public List<Span> Headers { get; } = new List<Span>();

            public override void EnterModuleHeader(VBAParser.ModuleHeaderContext context) => Enter(context);
            public override void ExitModuleHeader(VBAParser.ModuleHeaderContext context) => _inHeader--;
            public override void EnterModuleConfig(VBAParser.ModuleConfigContext context) => Enter(context);
            public override void ExitModuleConfig(VBAParser.ModuleConfigContext context) => _inHeader--;
            public override void EnterModuleConfigReferences(VBAParser.ModuleConfigReferencesContext context) => Enter(context);
            public override void ExitModuleConfigReferences(VBAParser.ModuleConfigReferencesContext context) => _inHeader--;

            private void Enter(ParserRuleContext context)
            {
                if (_inHeader++ == 0)
                {
                    Headers.Add(Of(context));
                }
            }

            public override void EnterCommentOrAnnotation(VBAParser.CommentOrAnnotationContext context)
            {
                //The context ends with the line break, which stays.
                if (_inHeader == 0 && context.GetChild(0) is ParserRuleContext comment)
                {
                    Spans.Add(Of(comment));
                }
            }
        }

        private class HeaderFinder : VBAParserBaseListener
        {
            public List<Span> Spans { get; } = new List<Span>();

            public override void EnterModuleHeader(VBAParser.ModuleHeaderContext context) => Spans.Add(Of(context));
            public override void EnterModuleConfigReferences(VBAParser.ModuleConfigReferencesContext context) => Spans.Add(Of(context));
            public override void EnterModuleConfig(VBAParser.ModuleConfigContext context)
            {
                if (!(context.Parent is VBAParser.ModuleConfigContext))
                {
                    Spans.Add(Of(context));
                }
            }

            public override void EnterAttributeStmt(VBAParser.AttributeStmtContext context) => Spans.Add(Of(context));
        }

        private class LineNumberFinder : VBAParserBaseListener
        {
            private readonly IList<IToken> _tokens;
            private readonly SortedSet<int> _outsideCodePane = new SortedSet<int>();
            private List<VBAParser.BlockStmtContext> _candidates;
            private bool _hasLineNumbers;

            public LineNumberFinder(IList<IToken> tokens)
            {
                _tokens = tokens;
            }

            public List<VBAParser.BlockStmtContext> Statements { get; } = new List<VBAParser.BlockStmtContext>();
            public List<string> Skipped { get; } = new List<string>();

            /// <summary>How many lines before this one the VBE does not show: the designer's header and every attribute line (also those after a procedure's signature).</summary>
            public int OutsideCodePaneBefore(int line) => _outsideCodePane.Count(l => l < line);

            public override void EnterModuleHeader(VBAParser.ModuleHeaderContext context) => Hide(context);
            public override void EnterModuleConfigReferences(VBAParser.ModuleConfigReferencesContext context) => Hide(context);
            public override void EnterModuleConfig(VBAParser.ModuleConfigContext context)
            {
                if (!(context.Parent is VBAParser.ModuleConfigContext))
                {
                    Hide(context);
                }
            }

            public override void EnterAttributeStmt(VBAParser.AttributeStmtContext context) => Hide(context);

            private void Hide(ParserRuleContext context)
            {
                if (context.Start == null || context.Stop == null)
                {
                    return;
                }

                for (var line = context.Start.Line; line <= context.Stop.Line; line++)
                {
                    _outsideCodePane.Add(line);
                }
            }

            public override void EnterSubStmt(VBAParser.SubStmtContext context) => Begin();
            public override void ExitSubStmt(VBAParser.SubStmtContext context) => End(context.subroutineName().GetText());
            public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context) => Begin();
            public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context) => End(context.functionName().GetText());
            public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context) => Begin();
            public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context) => End(context.functionName().GetText());
            public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context) => Begin();
            public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context) => End(context.subroutineName().GetText());
            public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context) => Begin();
            public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context) => End(context.subroutineName().GetText());

            private void Begin()
            {
                _candidates = new List<VBAParser.BlockStmtContext>();
                _hasLineNumbers = false;
            }

            private void End(string name)
            {
                if (_hasLineNumbers)
                {
                    Skipped.Add(name);
                }
                else
                {
                    Statements.AddRange(_candidates);
                }

                _candidates = null;
            }

            public override void EnterLineNumberLabel(VBAParser.LineNumberLabelContext context) => _hasLineNumbers = true;

            public override void EnterBlockStmt(VBAParser.BlockStmtContext context)
            {
                if (_candidates != null && IsNumberable(context))
                {
                    _candidates.Add(context);
                }
            }

            private bool IsNumberable(VBAParser.BlockStmtContext context)
            {
                var statement = context.mainBlockStmt();
                if (statement == null || context.statementLabelDefinition() != null
                    || statement.constStmt() != null || statement.variableStmt() != null || statement.attributeStmt() != null)
                {
                    return false;
                }

                //Only a statement that is the first thing on its physical line can carry a number.
                for (var i = context.Start.TokenIndex - 1; i >= 0; i--)
                {
                    if (_tokens[i].Type != VBALexer.WS)
                    {
                        return _tokens[i].Type == VBALexer.NEWLINE;
                    }
                }

                return true;
            }
        }

        private class AttributeFinder : VBAParserBaseListener
        {
            private static readonly string[] ModuleLevel =
            {
                "VB_Name", "VB_PredeclaredId", "VB_Exposed", "VB_Creatable", "VB_GlobalNameSpace", "VB_UserMemId"
            };

            private readonly ComponentType _type;
            private readonly Dictionary<string, bool> _publicProcedures = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            private readonly List<VBAParser.AttributeStmtContext> _attributes = new List<VBAParser.AttributeStmtContext>();
            private bool _privateModule;

            public AttributeFinder(ComponentType type)
            {
                _type = type;
            }

            public List<Span> Spans => _attributes.Where(a => !Keeps(a)).Select(Of).ToList();

            public override void EnterOptionPrivateModuleStmt(VBAParser.OptionPrivateModuleStmtContext context) => _privateModule = true;
            public override void EnterAttributeStmt(VBAParser.AttributeStmtContext context) => _attributes.Add(context);

            public override void EnterSubStmt(VBAParser.SubStmtContext context) => Procedure(context.subroutineName().GetText(), context.visibility());
            public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context) => Procedure(context.functionName().GetText(), context.visibility());
            public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context) => Procedure(context.functionName().GetText(), context.visibility());
            public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context) => Procedure(context.subroutineName().GetText(), context.visibility());
            public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context) => Procedure(context.subroutineName().GetText(), context.visibility());

            private void Procedure(string name, VBAParser.VisibilityContext visibility)
            {
                var text = visibility?.GetText();
                var isPublic = text == null
                               || text.Equals("Public", StringComparison.OrdinalIgnoreCase)
                               || text.Equals("Global", StringComparison.OrdinalIgnoreCase);
                _publicProcedures.TryGetValue(name.Trim('[', ']'), out var already);
                _publicProcedures[name.Trim('[', ']')] = already || isPublic;
            }

            private bool Keeps(VBAParser.AttributeStmtContext attribute)
            {
                var parts = attribute.attributeName().GetText().Trim().Split('.');
                bool Is(string part, string expected) => part.Equals(expected, StringComparison.OrdinalIgnoreCase);

                switch (parts.Length)
                {
                    case 1:
                        return ModuleLevel.Any(name => Is(parts[0], name));
                    //Default member, enumerator and hidden-member flags change how the code runs.
                    case 2 when Is(parts[1], "VB_UserMemId") || Is(parts[1], "VB_VarUserMemId") || Is(parts[1], "VB_MemberFlags"):
                        return true;
                    case 2 when Is(parts[1], "VB_Description"):
                        return ShownInExcel(parts[0]);
                    case 3 when Is(parts[1], "VB_ProcData")
                                && (Is(parts[2], "VB_Invoke_Func") || Is(parts[2], "VB_Invoke_PropertyPut") || Is(parts[2], "VB_Invoke_PropertyPutRef")):
                        return ShownInExcel(parts[0]);
                    default:
                        return false;
                }
            }

            private bool ShownInExcel(string procedure)
                => _type == ComponentType.StandardModule
                   && !_privateModule
                   && _publicProcedures.TryGetValue(procedure.Trim('[', ']'), out var isPublic)
                   && isPublic;
        }
    }
}
