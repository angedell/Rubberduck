using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Parsing.VBA.Parsing.ParsingExceptions;

namespace Rubberduck.Ship
{
    /// <summary>What <see cref="ShipDeclarationRemover"/> did to a module's text.</summary>
    public class ShipRemovalResult
    {
        public string Text { get; set; }
        public bool Parsed { get; set; } = true;
        public string Problem { get; set; }
        /// <summary>The removals that were applied (one per name).</summary>
        public List<ShipRemoval> Applied { get; } = new List<ShipRemoval>();
        /// <summary>Removals that were asked for and could not be applied, with why.</summary>
        public List<string> NotApplied { get; } = new List<string>();
        /// <summary>After the removals nothing is left in the module that is code: no procedure, variable, constant, Type, Enum or Declare. Option, Attribute, Event and Implements lines may be.</summary>
        public bool NothingLeft { get; set; }
    }

    /// <summary>
    /// Takes declarations out of the exported text of a module by name, with the spans of the parser's own parse tree (the module's text has a header and attribute lines the code pane does not,
    /// so the code pane's line numbers are never used). A procedure with its Get/Let/Set partners and the comment lines directly above it, a variable or constant statement, a Type, an Enum or a Declare goes with the attribute lines
    /// that belong to it and with the blank lines that separated it from what follows, so the module keeps the spacing it had. The result is parsed again; if it does not parse, the original text is returned.
    /// </summary>
    public class ShipDeclarationRemover
    {
        private readonly IStringParser _parser;

        public ShipDeclarationRemover(IStringParser parser)
        {
            _parser = parser;
        }

        private class Entry
        {
            public string Kind;
            public List<string> Names = new List<string>();
            public ParserRuleContext Context;
        }

        /// <summary>What the parse tree of a module says about it: the declarations, and what the removal needs to know about the lines around them.</summary>
        private class ParsedModule
        {
            public List<Entry> Entries { get; } = new List<Entry>();
            /// <summary>Where each comment (a ' comment, a Rem comment or an annotation list) starts, and whether an annotation in it is one of a module.</summary>
            public List<(int Start, bool ModuleAnnotation)> Comments { get; } = new List<(int, bool)>();
            /// <summary>Where each statement that is code starts: Option, Declare, Dim, Sub...; not the header and not the Attribute lines.</summary>
            public List<int> CodeStarts { get; } = new List<int>();
            /// <summary>The Attribute statements: the member they name (before the dot), the attribute (after it), and their span.</summary>
            public List<(string Member, string Attribute, int Start, int Stop)> Attributes { get; } = new List<(string, string, int, int)>();
        }

        //The annotations of the parser that apply to a module (AnnotationTarget.Module): '@Folder, '@ModuleDescription, '@TestModule...
        private static readonly HashSet<string> ModuleAnnotationNames = FindModuleAnnotationNames();

        private static HashSet<string> FindModuleAnnotationNames()
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var type in typeof(Rubberduck.Parsing.Annotations.IAnnotation).Assembly.GetTypes())
            {
                if (type.IsAbstract || !typeof(Rubberduck.Parsing.Annotations.IAnnotation).IsAssignableFrom(type) || type.GetConstructor(Type.EmptyTypes) == null)
                {
                    continue;
                }

                try
                {
                    var annotation = (Rubberduck.Parsing.Annotations.IAnnotation)Activator.CreateInstance(type);
                    if (annotation.Target.HasFlag(Rubberduck.Parsing.Annotations.AnnotationTarget.Module))
                    {
                        names.Add(annotation.Name);
                    }
                }
                catch (Exception)
                {
                    //An annotation that cannot be made here is not known as a module annotation; the fallback is that its comment goes with the member.
                }
            }

            return names;
        }

        private class Collector : VBAParserBaseListener
        {
            public ParsedModule Module { get; } = new ParsedModule();
            public List<Entry> Entries => Module.Entries;

            public override void EnterComment(VBAParser.CommentContext context) => Module.Comments.Add((context.Start.StartIndex, false));
            public override void EnterRemComment(VBAParser.RemCommentContext context) => Module.Comments.Add((context.Start.StartIndex, false));

            public override void EnterAnnotationList(VBAParser.AnnotationListContext context)
                => Module.Comments.Add((context.Start.StartIndex, context.annotation().Any(a => ModuleAnnotationNames.Contains(a.annotationName().GetText()))));

            public override void EnterModuleDeclarationsElement(VBAParser.ModuleDeclarationsElementContext context)
            {
                if (context.attributeStmt() == null)
                {
                    Module.CodeStarts.Add(context.Start.StartIndex);
                }
            }

            public override void EnterModuleBodyElement(VBAParser.ModuleBodyElementContext context) => Module.CodeStarts.Add(context.Start.StartIndex);

            public override void EnterAttributeStmt(VBAParser.AttributeStmtContext context)
            {
                var name = context.attributeName().GetText();
                var dot = name.IndexOf('.');
                Module.Attributes.Add((dot < 0 ? string.Empty : Plain(name.Substring(0, dot)), dot < 0 ? name : name.Substring(dot + 1), context.Start.StartIndex, context.Stop.StopIndex));
            }

            private void Add(string kind, ParserRuleContext context, params string[] names)
            {
                var entry = new Entry { Kind = kind, Context = context };
                entry.Names.AddRange(names);
                Entries.Add(entry);
            }

            public override void EnterSubStmt(VBAParser.SubStmtContext context) => Add("procedure", context, Plain(context.subroutineName().GetText()));
            public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context) => Add("procedure", context, Plain(context.functionName().GetText()));
            public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context) => Add("procedure", context, Plain(context.functionName().GetText()));
            public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context) => Add("procedure", context, Plain(context.subroutineName().GetText()));
            public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context) => Add("procedure", context, Plain(context.subroutineName().GetText()));

            public override void EnterModuleVariableStmt(VBAParser.ModuleVariableStmtContext context)
                => Add("variable", context, context.variableStmt().variableListStmt().variableSubStmt().Select(v => Plain(v.identifier().GetText())).ToArray());

            public override void EnterModuleConstStmt(VBAParser.ModuleConstStmtContext context)
                => Add("constant", context, context.constStmt().constSubStmt().Select(c => Plain(c.identifier().GetText())).ToArray());

            public override void EnterUdtDeclaration(VBAParser.UdtDeclarationContext context) => Add("type", context, Plain(context.untypedIdentifier().GetText()));
            public override void EnterEnumerationStmt(VBAParser.EnumerationStmtContext context) => Add("enum", context, Plain(context.identifier().GetText()));
            public override void EnterDeclareStmt(VBAParser.DeclareStmtContext context) => Add("declare", context, Plain(context.identifier().GetText()));

            //Brackets and a trailing type hint are not part of the name.
            private static string Plain(string identifier)
            {
                identifier = identifier.Trim();
                if (identifier.Length > 1 && identifier[0] == '[' && identifier[identifier.Length - 1] == ']')
                {
                    return identifier.Substring(1, identifier.Length - 2);
                }

                return identifier.Length > 1 && "%&^!#@$".IndexOf(identifier[identifier.Length - 1]) >= 0 ? identifier.Substring(0, identifier.Length - 1) : identifier;
            }
        }

        private ParsedModule Parse(string moduleName, string projectId, string text, out string problem)
        {
            problem = null;
            try
            {
                var (tree, _) = _parser.Parse(moduleName, projectId, text, CancellationToken.None, CodeKind.AttributesCode);
                var collector = new Collector();
                ParseTreeWalker.Default.Walk(collector, tree);
                return collector.Module;
            }
            catch (SyntaxErrorException exception)
            {
                problem = $"line {exception.LineNumber}, column {exception.Position}: {(exception.Message ?? string.Empty).Split('\r', '\n')[0]}";
                return null;
            }
            catch (Exception exception) when (!(exception is OperationCanceledException))
            {
                problem = exception.Message;
                return null;
            }
        }

        public ShipRemovalResult Remove(string moduleName, string projectId, string text, IReadOnlyCollection<ShipRemoval> removals)
        {
            text = text ?? string.Empty;
            var result = new ShipRemovalResult { Text = text };
            var parsed = Parse(moduleName, projectId, text, out var problem);
            if (parsed == null)
            {
                result.Parsed = false;
                result.Problem = problem;
                return result;
            }

            bool Asked(string kind, string name) => removals.Any(r => r.Kind == kind && string.Equals(r.Name, name, StringComparison.OrdinalIgnoreCase));

            var chosen = new List<Entry>();
            foreach (var entry in parsed.Entries)
            {
                if (entry.Names.Count > 0 && entry.Names.All(n => Asked(entry.Kind, n)))
                {
                    if (OwnsItsLines(text, entry))
                    {
                        chosen.Add(entry);
                    }
                    else
                    {
                        result.NotApplied.Add($"{moduleName}.{string.Join(", ", entry.Names)}: shares its line with other statements, so it was kept");
                    }
                }
                else if (entry.Names.Any(n => Asked(entry.Kind, n)))
                {
                    result.NotApplied.Add($"{moduleName}.{string.Join(", ", entry.Names)}: only some of the names in this {entry.Kind} statement are unreachable, so it was kept");
                }
            }

            foreach (var removal in removals.Where(r => r.Kind != "module"))
            {
                if (!chosen.Any(e => e.Kind == removal.Kind && e.Names.Any(n => string.Equals(n, removal.Name, StringComparison.OrdinalIgnoreCase)))
                    && !result.NotApplied.Any(n => n.StartsWith(removal.Module + "." + removal.Name + ":", StringComparison.OrdinalIgnoreCase)))
                {
                    //Not found in the text: it is in a conditional-compilation branch that is not active, or it was already gone.
                    result.NotApplied.Add($"{removal}: not found in the module's text");
                }
            }

            var shaken = chosen.Count == 0 ? text : Delete(text, chosen, parsed);
            var after = chosen.Count == 0 ? parsed : Parse(moduleName, projectId, shaken, out problem);
            if (after == null)
            {
                result.Parsed = false;
                result.Problem = "the module does not parse after the removals: " + problem;
                return result;
            }

            result.Text = shaken;
            foreach (var entry in chosen)
            {
                foreach (var name in entry.Names.Where(n => !result.Applied.Any(a => a.Kind == entry.Kind && string.Equals(a.Name, n, StringComparison.OrdinalIgnoreCase))))
                {
                    result.Applied.Add(new ShipRemoval { Module = moduleName, Name = name, Kind = entry.Kind });
                }
            }

            result.NothingLeft = after.Entries.Count == 0;
            return result;
        }

        private static bool IsBlank(string line) => line.Trim('\r', '\n', ' ', '\t').Length == 0;

        //Only an entry that has its lines to itself can be cut out by line: a statement that shares a line with others (after a colon) stays.
        private static bool OwnsItsLines(string text, Entry entry)
        {
            var first = entry.Context.Start.StartIndex;
            var last = entry.Context.Stop.StopIndex;
            var lineStart = first == 0 ? 0 : text.LastIndexOf('\n', first - 1) + 1;
            var lineEnd = text.IndexOf('\n', last + 1);
            var before = text.Substring(lineStart, first - lineStart);
            var rest = text.Substring(last + 1, (lineEnd < 0 ? text.Length : lineEnd) - last - 1).Trim();
            return before.Trim().Length == 0 && (rest.Length == 0 || rest[0] == '\'' || rest.StartsWith("Rem ", StringComparison.OrdinalIgnoreCase));
        }

        //Deletes the lines of the entries. Where the removal joins stretches of blank lines (before and after the entry, between entries that follow each other),
        //the longest of them stays, so the module keeps the spacing between its parts and no longer run of blank lines appears. At the start and at the end of the module none stay.
        private static string Delete(string text, List<Entry> chosen, ParsedModule parsed)
        {
            var lines = SplitLines(text);
            var offsets = new int[lines.Count + 1];
            for (var i = 0; i < lines.Count; i++)
            {
                offsets[i + 1] = offsets[i] + lines[i].Length;
            }

            int LineOf(int index)
            {
                var low = 0;
                var high = lines.Count - 1;
                while (low < high)
                {
                    var middle = (low + high + 1) / 2;
                    if (offsets[middle] <= index)
                    {
                        low = middle;
                    }
                    else
                    {
                        high = middle - 1;
                    }
                }

                return low;
            }

            //The lines that hold nothing but a comment, found with the comment contexts of the parse tree.
            var commentLine = new bool[lines.Count];
            var moduleAnnotationLine = new bool[lines.Count];
            foreach (var (start, moduleAnnotation) in parsed.Comments)
            {
                var line = LineOf(start);
                if (text.Substring(offsets[line], start - offsets[line]).Trim().Length == 0)
                {
                    commentLine[line] = true;
                    moduleAnnotationLine[line] |= moduleAnnotation;
                }
            }

            var drop = new bool[lines.Count];
            foreach (var entry in chosen)
            {
                var first = LineOf(entry.Context.Start.StartIndex);
                for (var line = first; line <= LineOf(entry.Context.Stop.StopIndex); line++)
                {
                    drop[line] = true;
                }

                //The comment and annotation lines directly above (no blank line between) describe this declaration and go with it. Not a module annotation
                //('@Folder...), and not the comment at the very top of the code, which describes the module.
                var top = first;
                while (top - 1 >= 0 && commentLine[top - 1] && !moduleAnnotationLine[top - 1])
                {
                    top--;
                }

                if (top < first && parsed.CodeStarts.Any(start => start < offsets[top]))
                {
                    for (var line = top; line < first; line++)
                    {
                        drop[line] = true;
                    }
                }
            }

            //Attribute mApp.VB_VarHelpID = -1 and the like follow the variable they belong to; without it they would name a member that is gone.
            var variables = new HashSet<string>(chosen.Where(e => e.Kind == "variable" || e.Kind == "constant").SelectMany(e => e.Names), StringComparer.OrdinalIgnoreCase);
            foreach (var attribute in parsed.Attributes.Where(x => variables.Contains(x.Member) && x.Attribute.StartsWith("VB_Var", StringComparison.OrdinalIgnoreCase)))
            {
                for (var line = LineOf(attribute.Start); line <= LineOf(attribute.Stop); line++)
                {
                    drop[line] = true;
                }
            }

            var builder = new StringBuilder(text.Length);
            var index2 = 0;
            while (index2 < lines.Count)
            {
                var skip = drop[index2];
                if (!skip && !IsBlank(lines[index2]))
                {
                    builder.Append(lines[index2++]);
                    continue;
                }

                //A stretch of blank and dropped lines.
                var start = index2;
                var touched = false;
                var longest = 0;
                var current = 0;
                string blank = null;
                while (index2 < lines.Count && (drop[index2] || IsBlank(lines[index2])))
                {
                    if (drop[index2])
                    {
                        touched = true;
                        current = 0;
                    }
                    else
                    {
                        current++;
                        longest = Math.Max(longest, current);
                        blank = blank ?? lines[index2];
                    }

                    index2++;
                }

                if (!touched)
                {
                    for (var i = start; i < index2; i++)
                    {
                        builder.Append(lines[i]);
                    }

                    continue;
                }

                var atStart = start == 0 || lines.Take(start).All(IsBlank);
                var atEnd = index2 >= lines.Count;
                for (var i = 0; !atStart && !atEnd && i < longest; i++)
                {
                    builder.Append(blank);
                }
            }

            return builder.ToString();
        }

        //Lines with their line breaks.
        private static List<string> SplitLines(string text)
        {
            var lines = new List<string>();
            var start = 0;
            for (var i = 0; i < text.Length; i++)
            {
                if (text[i] == '\n')
                {
                    lines.Add(text.Substring(start, i - start + 1));
                    start = i + 1;
                }
            }

            if (start < text.Length)
            {
                lines.Add(text.Substring(start));
            }

            return lines;
        }
    }
}
