using System;
using System.Collections.Generic;
using Stream = System.IO.Stream;
using System.Linq;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Ship
{
    /// <summary>A module of the project as exported text, and what the steps did to it. The user's own module is never touched: this is a copy of its text.</summary>
    public class ShipModule
    {
        public ShipModule(string name, ComponentType type, string text)
        {
            Name = name;
            Type = type;
            OriginalText = text;
            Text = text;
        }

        public string Name { get; }
        public ComponentType Type { get; }

        /// <summary>The whole exported file: header, attributes and code.</summary>
        public string OriginalText { get; }

        /// <summary>The file the module was exported to (read-only copy of the user's module); a form's .frx is next to it.</summary>
        public string ExportFile { get; set; }

        /// <summary>The text after the steps so far.</summary>
        public string Text { get; set; }

        /// <summary>The module is not part of the artifact.</summary>
        public bool Removed { get; set; }

        /// <summary>Set when a step could not parse the module; it ships as it was.</summary>
        public string UnparsedReason { get; set; }

        public bool Changed => !Removed && !string.Equals(Text, OriginalText, StringComparison.Ordinal);
    }

    public class ShipReference
    {
        public string Name { get; set; }
        public string Guid { get; set; }
        public int Major { get; set; }
        public int Minor { get; set; }
        public bool Removed { get; set; }
    }

    /// <summary>
    /// What the strip steps work on: the project's modules as text, its references, and the parts of the document's file. Steps change this, never the project;
    /// <see cref="ShipArtifactWriter"/> turns the result into the artifact.
    /// </summary>
    public class ShipWorkingSet
    {
        public ShipWorkingSet(string projectId)
        {
            ProjectId = projectId;
        }

        public string ProjectId { get; }
        public List<ShipModule> Modules { get; } = new List<ShipModule>();
        public List<ShipReference> References { get; } = new List<ShipReference>();

        /// <summary>Opens the saved document's file for reading (it is a zip of parts); null when the file is not available to the steps.</summary>
        public Func<Stream> OpenPackage { get; set; }

        /// <summary>Names of the parts (as <c>PackageResource.Name</c>) to take out of the artifact.</summary>
        public List<string> RemovedResources { get; } = new List<string>();

        public List<string> Warnings { get; } = new List<string>();
    }

    /// <summary>One kind of removal. Every step is independent of the others and says what it removed.</summary>
    public interface IShipStep
    {
        /// <summary>One of <see cref="ShipSteps"/>.</summary>
        string Name { get; }

        ShipStepReport Run(ShipWorkingSet set);
    }

    /// <summary>Runs the enabled steps in order over a working set and builds the report. It writes nothing; a dry run is the same run, marked.</summary>
    public class ShipPipeline
    {
        private readonly IReadOnlyList<IShipStep> _steps;

        public ShipPipeline(IStringParser parser)
        {
            var text = new ShipModuleText(parser);
            _steps = new IShipStep[]
            {
                new RemoveNotesStep(),
                new RemoveTestModulesStep(),
                new RemoveCommentsStep(text),
                new RemoveDescriptiveAttributesStep(text),
                new RemoveConfigStep()
            };
        }

        public ShipReport Run(ShipWorkingSet set, ShipSettings settings, bool dryRun)
        {
            var report = new ShipReport { DryRun = dryRun };
            foreach (var step in _steps.Where(s => settings.IsOn(s.Name)))
            {
                report.Steps.Add(step.Run(set));
            }

            report.Unparsed.AddRange(set.Modules.Where(m => !m.Removed && m.UnparsedReason != null).Select(m => $"{m.Name}: {m.UnparsedReason}"));
            report.Warnings.AddRange(set.Warnings);
            return report;
        }
    }
}
