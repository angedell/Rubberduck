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
        public const string Comments = "comments";
        public const string DescriptiveAttributes = "descriptiveAttributes";
        public const string Config = "config";

        public static readonly string[] All = { Notes, TestModules, Comments, DescriptiveAttributes, Config };
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
        public bool IsOn(string step) => !Steps.TryGetValue(step, out var on) || on;
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
