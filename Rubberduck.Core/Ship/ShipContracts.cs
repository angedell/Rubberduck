using System;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Ship
{
    public enum ShipFormat { Xlsm, Xlam }

    /// <summary>What to do when the artifact's file exists. <see cref="Ask"/> needs an answer from the caller.</summary>
    public enum ShipOnExisting { Ask, Replace, KeepBoth }

    /// <summary>The names of the strip steps, as they appear in the settings and in reports.</summary>
    public static class ShipSteps
    {
        public const string Notes = "notes";
        public const string TestModules = "testModules";
        public const string LineNumbers = "lineNumbers";
        public const string TreeShaking = "treeShaking";
        public const string Comments = "comments";
        public const string DescriptiveAttributes = "descriptiveAttributes";
        public const string Config = "config";

        public static readonly string[] All = { Notes, TestModules, LineNumbers, TreeShaking, Comments, DescriptiveAttributes, Config };

        /// <summary>Steps that are off until the user switches them on: tree shaking can remove code the user needs, line numbers change every line of the code that ships.</summary>
        public static bool OffByDefault(string step) => step == TreeShaking || step == LineNumbers;
    }

    /// <summary>What is a root of tree shaking besides the always-kept code (see ADR 0004).</summary>
    public enum ShipTreeShakingMode
    {
        /// <summary>Every public procedure of a standard module is a root: it may be a macro or a UDF.</summary>
        Conservative,
        /// <summary>Only @EntryPoint / @ExcelHotkey members, the keep list and the detected roots.</summary>
        EntryPointsOnly
    }

    public class ShipTreeShakingSettings
    {
        public ShipTreeShakingMode Mode { get; set; } = ShipTreeShakingMode.Conservative;
        /// <summary>Names to keep whatever the analysis says: "Member" or "Module.Member" (case-insensitive).</summary>
        public List<string> Keep { get; } = new List<string>();
    }

    /// <summary>The <c>ship</c> section of the project's rubberduck.json. Null members mean "not set in the file".</summary>
    public class ShipSettings
    {
        public ShipFormat Format { get; set; } = ShipFormat.Xlsm;
        /// <summary>As written in the file: may be relative, use ../, %ENV%, or a known/OneDrive/Dropbox folder.</summary>
        public string OutputFolder { get; set; }
        /// <summary>Without extension; null means the project's name.</summary>
        public string FileName { get; set; }
        public ShipOnExisting OnExisting { get; set; } = ShipOnExisting.Ask;
        /// <summary>Step name (see <see cref="ShipSteps"/>) to on/off. A step that is not listed is on.</summary>
        public Dictionary<string, bool> Steps { get; } = new Dictionary<string, bool>();
        public ShipTreeShakingSettings TreeShaking { get; } = new ShipTreeShakingSettings();
        public bool IsOn(string step) => Steps.TryGetValue(step, out var on) ? on : !ShipSteps.OffByDefault(step);
    }

    public class ShipStepReport
    {
        public string Step { get; set; }
        public int Removed { get; set; }
        /// <summary>Names of what was removed (modules, references, parts) or a count line, for display.</summary>
        public List<string> Details { get; } = new List<string>();
    }

    public class ShipReport
    {
        public List<ShipStepReport> Steps { get; } = new List<ShipStepReport>();
        /// <summary>Modules shipped unchanged because they did not parse, with why.</summary>
        public List<string> Unparsed { get; } = new List<string>();
        public List<string> Warnings { get; } = new List<string>();
        /// <summary>Calls whose target is computed (Application.Run, CallByName, OnTime with a non-literal name): the analysis cannot see what they reach. One line each.</summary>
        public List<string> DynamicCalls { get; } = new List<string>();
        /// <summary>Tree shaking: why each root is one ("Module1.Go: public procedure of a standard module"). Empty when the step did not run.</summary>
        public List<string> Roots { get; } = new List<string>();
        public string OutputPath { get; set; }
        /// <summary>True for a preview: nothing was written.</summary>
        public bool DryRun { get; set; }
    }

    /// <summary>Why a ship cannot start. Thrown for rules, never for bugs.</summary>
    public class ShipRefusedException : Exception
    {
        public ShipRefusedException(string reason, string message) : base(message) { Reason = reason; }
        /// <summary>unsupportedHost | unsaved | notExcel | folderMissing | folderNotWritable | exists | excelUnavailable | vbomNotTrusted</summary>
        public string Reason { get; }
        /// <summary>For <c>exists</c>: the path that is in the way.</summary>
        public string Path { get; set; }
    }

    /// <summary>One declaration tree shaking takes out of the artifact. Property Get/Let/Set of a name are one removal.</summary>
    public class ShipRemoval
    {
        public string Module { get; set; }
        public string Name { get; set; }
        /// <summary>procedure | variable | constant | type | enum | declare | module</summary>
        public string Kind { get; set; }
        public override string ToString() => Kind == "module" ? Module : Module + "." + Name;
    }

    /// <summary>
    /// The stage before the strip steps (ADR 0003 and 0004): reads what Rubberduck's parser resolved for the open project and says what is unreachable.
    /// It runs while the annotations are still in the text, and never changes anything.
    /// </summary>
    public interface IShipAnalysis
    {
        /// <summary>
        /// Null, with <paramref name="whyNot"/> set, when the parser is not ready or a module does not parse: tree shaking is then skipped as a whole.
        /// The result also carries the roots with their sources and the dynamic calls.
        /// </summary>
        ShipAnalysisResult Analyze(IVBProject project, ShipWorkingSet set, ShipTreeShakingSettings settings, out string whyNot);
    }

    public class ShipAnalysisResult
    {
        public List<ShipRemoval> Removals { get; } = new List<ShipRemoval>();
        /// <summary>Why each root is one: "Module1.Go: public procedure", "Foo: string literal in Module2", "Bar: ribbon onAction"...</summary>
        public List<string> Roots { get; } = new List<string>();
        public List<string> DynamicCalls { get; } = new List<string>();
        /// <summary>For the report: for example a keep entry that matches nothing.</summary>
        public List<string> Warnings { get; } = new List<string>();
    }

    public interface IShipService
    {
        /// <summary>Whether Ship is offered for this project at all (Excel document with a file).</summary>
        bool CanShip(IVBProject project);

        ShipSettings ReadSettings(IVBProject project);
        void WriteSettings(IVBProject project, ShipSettings settings);

        /// <summary>The folder the settings resolve to on this machine, or null if none is set.</summary>
        string ResolveOutputFolder(IVBProject project, ShipSettings settings);

        /// <summary>Runs the pipeline on the copy without saving anything. Throws <see cref="ShipRefusedException"/>.</summary>
        ShipReport Preview(IVBProject project, ShipSettings settings);

        /// <summary>Writes the artifact. <paramref name="onExisting"/> overrides the settings if it is not <see cref="ShipOnExisting.Ask"/>.</summary>
        ShipReport Ship(IVBProject project, ShipSettings settings, ShipOnExisting onExisting, IProgress<string> progress);

        /// <summary>The user's own switch (see glossary: unattended ship). Setting is for the dialog only.</summary>
        bool IsUnattended(IVBProject project);
        void SetUnattended(IVBProject project, bool value);
    }
}
