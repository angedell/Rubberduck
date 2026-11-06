using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Ship;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using static RubberduckTests.Ship.ShipTestParser;

namespace RubberduckTests.Ship
{
    /// <summary>A module of a small project, as the code pane holds it and as the VBE exports it.</summary>
    internal class CorpusModule
    {
        public string Name { get; set; }
        public ComponentType Type { get; set; }
        public string Code { get; set; }
        public bool Exposed { get; set; }

        public static CorpusModule Std(string name, params string[] lines) => new CorpusModule { Name = name, Type = ComponentType.StandardModule, Code = Lines(lines) };
        public static CorpusModule Class(string name, params string[] lines) => new CorpusModule { Name = name, Type = ComponentType.ClassModule, Code = Lines(lines) };
        public static CorpusModule ExposedClass(string name, params string[] lines) => new CorpusModule { Name = name, Type = ComponentType.ClassModule, Code = Lines(lines), Exposed = true };
        public static CorpusModule Document(string name, params string[] lines) => new CorpusModule { Name = name, Type = ComponentType.Document, Code = Lines(lines), Exposed = true };
        public static CorpusModule Form(string name, params string[] lines) => new CorpusModule { Name = name, Type = ComponentType.UserForm, Code = Lines(lines) };

        public CorpusModule WithCode(string code) => new CorpusModule { Name = Name, Type = Type, Code = code, Exposed = Exposed };

        /// <summary>What the mock code pane holds. The mock parser reads attributes from the code, so an exposed class says so there.</summary>
        public string CodePane => (Exposed && Type == ComponentType.ClassModule ? "Attribute VB_Exposed = True\r\n" : string.Empty) + Code;

        /// <summary>The file the VBE exports for this module: header, attributes, code.</summary>
        public string Exported
        {
            get
            {
                switch (Type)
                {
                    case ComponentType.StandardModule:
                        return $"Attribute VB_Name = \"{Name}\"\r\n{Code}";
                    case ComponentType.UserForm:
                        return $"VERSION 5.00\r\nBegin {{C62A69F0-16DC-11CE-9E98-00AA00574A4F}} {Name} \r\n   Caption         =   \"{Name}\"\r\n   ClientHeight    =   3015\r\n   OleObjectBlob   =   \"{Name}.frx\":0000\r\nEnd\r\n"
                               + $"Attribute VB_Name = \"{Name}\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = False\r\n{Code}";
                    default:
                        return "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n"
                               + $"Attribute VB_Name = \"{Name}\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\n"
                               + $"Attribute VB_PredeclaredId = {(Type == ComponentType.Document ? "True" : "False")}\r\nAttribute VB_Exposed = {(Exposed ? "True" : "False")}\r\n{Code}";
                }
        }
        }
    }

    /// <summary>A small project parsed for real (the parser Rubberduck uses, over a mock VBE): what tree shaking is tested on.</summary>
    internal class Corpus
    {
        public CorpusModule[] Modules { get; private set; }
        public RubberduckParserState State { get; private set; }
        public IVBProject Project { get; private set; }
        public string ProjectId => Project.ProjectId;

        public static Corpus Parse(params CorpusModule[] modules)
        {
            var vbe = MockVbeBuilder.BuildFromModules("TestProject1", modules.Select(m => (m.Name, m.CodePane, m.Type)), new[] { ReferenceLibrary.VBA, ReferenceLibrary.Excel, ReferenceLibrary.MsOffice });
            var parser = MockParser.Create(vbe.Object);
            parser.Parse(new CancellationTokenSource());
            return new Corpus { Modules = modules, State = parser.State, Project = vbe.Object.ActiveVBProject };
        }

        public ShipWorkingSet Set(Func<System.IO.Stream> package = null)
        {
            var set = new ShipWorkingSet(ProjectId) { OpenPackage = package };
            set.Modules.AddRange(Modules.Select(m => new ShipModule(m.Name, m.Type, m.Exported)));
            return set;
        }

        public ParserStateShipAnalysis Analysis() => new ParserStateShipAnalysis(State, CreateFor(null, ProjectId), () => State.Status);

        public ShipAnalysisResult Analyze(ShipTreeShakingSettings settings, out string whyNot, Func<System.IO.Stream> package = null)
            => Analysis().Analyze(Project, Set(package), settings ?? new ShipTreeShakingSettings(), out whyNot);

        /// <summary>The settings of a tree-shaking-only run.</summary>
        public static ShipSettings Only(ShipTreeShakingMode mode = ShipTreeShakingMode.Conservative, params string[] keep)
        {
            var settings = new ShipSettings();
            foreach (var step in ShipSteps.All)
            {
                settings.Steps[step] = step == ShipSteps.TreeShaking;
            }

            settings.TreeShaking.Mode = mode;
            settings.TreeShaking.Keep.AddRange(keep);
            return settings;
        }

        /// <summary>Runs the pipeline with tree shaking alone and returns the working set and the report.</summary>
        public (ShipWorkingSet Set, ShipReport Report) Shake(ShipSettings settings, Func<System.IO.Stream> package = null)
        {
            var set = Set(package);
            var pipeline = new ShipPipeline(CreateFor(null, ProjectId), Analysis());
            return (set, pipeline.Run(set, settings, dryRun: false, project: Project));
        }

        public static string[] Names(IEnumerable<ShipRemoval> removals) => removals.Select(r => r.ToString()).OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToArray();

        // ---- the invariant -------------------------------------------------------------------------------------------------------------

        /// <summary>
        /// The shaken text is parsed again, for real, as a project of its own, and it must (1) parse and resolve, (2) keep every reference that a surviving member had in this project,
        /// resolved to the same declaration, (3) have no declaration whose type does not resolve that did before, (4) no new undeclared identifier or unresolved member.
        /// </summary>
        public void AssertShakenIsSound(ShipWorkingSet shaken)
        {
            var text = new ShipModuleText(CreateFor(null, ProjectId));
            var kept = new List<CorpusModule>();
            foreach (var module in Modules)
            {
                var shakenModule = shaken.Modules.Single(m => m.Name == module.Name);
                if (shakenModule.Removed)
                {
                    continue;
                }

                var codeOnly = text.CodeOnly(module.Name, ProjectId, shakenModule.Text);
                Assert.IsTrue(codeOnly.Parsed, $"The shaken text of {module.Name} does not parse: {codeOnly.Problem}\r\n{shakenModule.Text}");
                //The mock parser reads member attributes (VB_UserMemId...) from the code, as the first parse did, so only the header of the exported text is dropped.
                kept.Add(module.WithCode(WithoutHeader(shakenModule.Text)));
            }

            if (kept.Count == 0)
            {
                return;
            }

            var after = Parse(kept.ToArray());
            Assert.AreEqual(ParserState.Ready, after.State.Status, "The shaken project did not parse and resolve cleanly:\r\n" + string.Join("\r\n", kept.Select(k => "--- " + k.Name + "\r\n" + k.Code)));

            var ownersAfter = new HashSet<string>(after.References().Select(r => r.Owner));
            var beforeReferences = References();
            var afterSignatures = new HashSet<string>(after.References().Select(r => r.Signature));
            var missing = beforeReferences.Where(r => ownersAfter.Contains(r.Owner) && !afterSignatures.Contains(r.Signature)).Select(r => r.Signature).Distinct().ToList();
            CollectionAssert.IsEmpty(missing, "A surviving member lost a reference:\r\n" + string.Join("\r\n", kept.Select(k => "--- " + k.Name + "\r\n" + k.Code)));

            var unresolvedBefore = Unresolved();
            var newlyUnresolved = after.Unresolved().Except(unresolvedBefore).ToList();
            CollectionAssert.IsEmpty(newlyUnresolved, "The shaken project has identifiers or types that do not resolve:\r\n" + string.Join("\r\n", kept.Select(k => "--- " + k.Name + "\r\n" + k.Code)));
        }

        private static readonly System.Text.RegularExpressions.Regex HeaderLine = new System.Text.RegularExpressions.Regex(
            @"^(VERSION\b|BEGIN\b|END\s*$|\s+\w+\s*=|Begin\b|Attribute VB_(Name|GlobalNameSpace|Creatable|PredeclaredId|Exposed|Customizable|TemplateDerived)\b)",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        private static string WithoutHeader(string exported)
        {
            var lines = exported.Split('\n').ToList();
            var skip = 0;
            while (skip < lines.Count && HeaderLine.IsMatch(lines[skip].TrimEnd('\r')))
            {
                skip++;
            }

            return string.Join("\n", lines.Skip(skip)).TrimStart('\r', '\n');
        }

        internal List<(string Owner, string Signature)> References()
        {
            var list = new List<(string Owner, string Signature)>();
            foreach (var declaration in State.AllUserDeclarations.Where(d => d.ProjectId == ProjectId))
            {
                foreach (var reference in declaration.References)
                {
                    var owner = reference.ParentScoping;
                    if (owner == null || owner.DeclarationType.HasFlag(DeclarationType.Module))
                    {
                        continue;
                    }

                    var ownerKey = owner.QualifiedModuleName.ComponentName + "." + owner.IdentifierName + ":" + owner.DeclarationType;
                    list.Add((ownerKey, $"{ownerKey}|{reference.IdentifierName.ToLowerInvariant()}|{declaration.QualifiedModuleName.ComponentName}.{declaration.IdentifierName}:{declaration.DeclarationType}"));
                }
            }

            return list;
        }

        internal HashSet<string> Unresolved()
        {
            var set = new HashSet<string>();
            foreach (var declaration in State.AllUserDeclarations.Where(d => d.ProjectId == ProjectId))
            {
                if (declaration.IsUndeclared)
                {
                    set.Add($"undeclared {declaration.QualifiedModuleName.ComponentName}.{declaration.IdentifierName}");
                }

                if (!string.IsNullOrEmpty(declaration.AsTypeName) && declaration.AsTypeDeclaration == null && !declaration.IsUndeclared
                    && declaration.DeclarationType != DeclarationType.UnresolvedMember && IsUserTypeName(declaration.AsTypeName))
                {
                    set.Add($"type {declaration.QualifiedModuleName.ComponentName}.{declaration.IdentifierName} As {declaration.AsTypeName}");
                }
            }

            foreach (var member in State.DeclarationFinder.UnresolvedMemberDeclarations().Where(u => u.ProjectId == ProjectId))
            {
                set.Add($"member {member.QualifiedModuleName.ComponentName}.{member.IdentifierName}");
            }

            return set;
        }

        //A type name that is not one of VBA's own; "Long", "String", "Object"... have no declaration to resolve to.
        private static readonly HashSet<string> BuiltInTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Boolean", "Byte", "Currency", "Date", "Double", "Integer", "Long", "LongLong", "LongPtr", "Object", "Single", "String", "Variant", "Any"
        };

        private static bool IsUserTypeName(string name) => !BuiltInTypes.Contains(name) && !name.Contains(".");
    }
}
