using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Ship
{
    /// <summary>
    /// Takes out what the analysis (<see cref="IShipAnalysis"/>) found unreachable, by name, with the parse tree's own spans of each module's exported text (see ADR 0004).
    /// User forms are never touched. A module whose removal is asked for and that has no code left is dropped from the artifact. A module that does not parse after the edit ships as it was.
    /// </summary>
    public class TreeShakingStep : IShipStep
    {
        private readonly ShipDeclarationRemover _remover;
        private readonly IReadOnlyList<ShipRemoval> _removals;

        public TreeShakingStep(ShipDeclarationRemover remover, IReadOnlyList<ShipRemoval> removals)
        {
            _remover = remover;
            _removals = removals ?? new List<ShipRemoval>();
        }

        public string Name => ShipSteps.TreeShaking;

        public ShipStepReport Run(ShipWorkingSet set)
        {
            var report = new ShipStepReport { Step = Name };

            //Nothing is changed until every module has been shaken: if one cannot be, the step is skipped as a whole and the set stays as it was.
            var notes = new List<string>();
            var plan = new List<(ShipModule module, bool dropModule, string text, ShipRemovalResult result)>();
            foreach (var module in set.Modules.Where(m => !m.Removed && m.UnparsedReason == null && m.Type != ComponentType.UserForm))
            {
                var mine = _removals.Where(r => string.Equals(r.Module, module.Name, StringComparison.OrdinalIgnoreCase)).ToList();
                if (mine.Count == 0)
                {
                    continue;
                }

                var wholeModule = mine.Any(r => r.Kind == "module") && (module.Type == ComponentType.StandardModule || module.Type == ComponentType.ClassModule);
                var result = _remover.Remove(module.Name, set.ProjectId, module.Text, mine.Where(r => r.Kind != "module").ToList());
                if (!result.Parsed)
                {
                    return Skipped(set, report, $"module {module.Name} could not be shaken: {result.Problem}");
                }

                if (result.NotApplied.Count > 0)
                {
                    return Skipped(set, report, $"module {module.Name} could not be shaken: {string.Join("; ", result.NotApplied)}");
                }

                var drop = wholeModule && result.NothingLeft;
                if (wholeModule && !drop)
                {
                    notes.Add($"Tree shaking: module {module.Name} is unreachable but still holds code that could not be taken out, so it was kept.");
                }

                plan.Add((module, drop, result.Text, result));
            }

            set.Warnings.AddRange(notes);
            var names = new List<string>();
            var modulesRemoved = 0;
            var declarations = 0;
            foreach (var (module, drop, text, result) in plan)
            {
                if (drop)
                {
                    module.Removed = true;
                    modulesRemoved++;
                    names.Add(new ShipRemoval { Module = module.Name, Name = module.Name, Kind = "module" }.ToString());
                    continue;
                }

                module.Text = text;
                declarations += result.Applied.Count;
                names.AddRange(result.Applied.Select(r => r.ToString()));
            }

            report.Removed = declarations + modulesRemoved;
            report.Details.Add($"{declarations} declaration(s) and {modulesRemoved} module(s)");
            report.Details.AddRange(names);
            return report;
        }

        private static ShipStepReport Skipped(ShipWorkingSet set, ShipStepReport report, string why)
        {
            set.Warnings.Add("Tree shaking was skipped, nothing was removed: " + why);
            report.Removed = 0;
            report.Details.Add("Skipped: " + why);
            return report;
        }
    }
}
