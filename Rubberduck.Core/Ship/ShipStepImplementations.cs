using System;
using System.Linq;
using System.Text.RegularExpressions;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.PackageResources;
using Rubberduck.Resources.Registration;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Ship
{
    /// <summary>Takes out the notes: standard modules that hold nothing but a @Notes marker and comments.</summary>
    public class RemoveNotesStep : IShipStep
    {
        public string Name => ShipSteps.Notes;

        public ShipStepReport Run(ShipWorkingSet set)
        {
            var report = new ShipStepReport { Step = Name };
            foreach (var module in set.Modules.Where(m => !m.Removed && m.Type == ComponentType.StandardModule && NoteModule.IsNote(m.Text)))
            {
                module.Removed = true;
                report.Removed++;
                report.Details.Add(NoteModule.DisplayName(module.Name, module.Text));
            }

            return report;
        }
    }

    /// <summary>Takes out the test modules (marked @TestModule) and the reference to Rubberduck's unit-testing library.</summary>
    public class RemoveTestModulesStep : IShipStep
    {
        private static readonly Regex TestModuleAnnotation = new Regex(@"^[ \t]*'@TestModule\b", RegexOptions.IgnoreCase | RegexOptions.Multiline);
        private static readonly Regex LibraryUse = new Regex(@"\bRubberduck\s*\.\s*(AssertClass|PermissiveAssertClass|FakesProvider|IAssert|IFakesProvider|IFake|IVerify)\b", RegexOptions.IgnoreCase);

        public string Name => ShipSteps.TestModules;

        public ShipStepReport Run(ShipWorkingSet set)
        {
            var report = new ShipStepReport { Step = Name };
            foreach (var module in set.Modules.Where(m => !m.Removed && TestModuleAnnotation.IsMatch(m.Text)))
            {
                module.Removed = true;
                report.Removed++;
                report.Details.Add(module.Name);
            }

            var library = RubberduckGuid.RubberduckTypeLibGuid.Trim('{', '}');
            foreach (var reference in set.References.Where(r => !r.Removed && string.Equals((r.Guid ?? string.Empty).Trim('{', '}'), library, StringComparison.OrdinalIgnoreCase)))
            {
                reference.Removed = true;
                report.Removed++;
                report.Details.Add("reference " + reference.Name);
            }

            if (set.References.Any(r => r.Removed && string.Equals((r.Guid ?? string.Empty).Trim('{', '}'), library, StringComparison.OrdinalIgnoreCase)))
            {
                foreach (var module in set.Modules.Where(m => !m.Removed && LibraryUse.IsMatch(m.Text)))
                {
                    set.Warnings.Add($"Module {module.Name} still uses Rubberduck's unit-testing library, whose reference is removed; it will not compile in the shipped project.");
                }
            }

            return report;
        }
    }

    /// <summary>
    /// Numbers the executable lines of every procedure with their line in the user's own module, so the shipped project's error handlers can report <c>Erl</c>.
    /// It runs before the steps that take lines out, which is what keeps the numbers true to the source.
    /// </summary>
    public class AddLineNumbersStep : IShipStep
    {
        private readonly ShipModuleText _text;

        public AddLineNumbersStep(ShipModuleText text)
        {
            _text = text;
        }

        public string Name => ShipSteps.LineNumbers;

        public ShipStepReport Run(ShipWorkingSet set)
        {
            var report = new ShipStepReport { Step = Name };
            var modules = 0;
            foreach (var module in set.Modules.Where(m => !m.Removed && m.UnparsedReason == null))
            {
                var result = _text.NumberLines(module.Name, set.ProjectId, module.Text);
                if (!result.Parsed)
                {
                    module.UnparsedReason = result.Problem;
                    continue;
                }

                module.Text = result.Text;
                if (result.Numbered > 0)
                {
                    modules++;
                    report.Removed += result.Numbered;
                }

                foreach (var procedure in result.Skipped)
                {
                    set.Warnings.Add($"Line numbers: {module.Name}.{procedure} has line numbers of its own and was left as it is.");
                }
            }

            report.Details.Add($"{report.Removed} line(s) in {modules} module(s)");
            return report;
        }
    }

    /// <summary>Takes out every comment and annotation, with the parser.</summary>
    public class RemoveCommentsStep : IShipStep
    {
        private readonly ShipModuleText _text;

        public RemoveCommentsStep(ShipModuleText text)
        {
            _text = text;
        }

        public string Name => ShipSteps.Comments;

        public ShipStepReport Run(ShipWorkingSet set)
        {
            var report = new ShipStepReport { Step = Name };
            var modules = 0;
            foreach (var module in set.Modules.Where(m => !m.Removed && m.UnparsedReason == null))
            {
                var result = _text.StripComments(module.Name, set.ProjectId, module.Text);
                if (!result.Parsed)
                {
                    module.UnparsedReason = result.Problem;
                    continue;
                }

                module.Text = result.Text;
                if (result.Removed > 0)
                {
                    modules++;
                    report.Removed += result.Removed;
                }
            }

            report.Details.Add($"{report.Removed} comment(s) in {modules} module(s)");
            return report;
        }
    }

    /// <summary>Takes out the attribute lines that only describe: see <see cref="ShipModuleText.StripDescriptiveAttributes"/> for the rule.</summary>
    public class RemoveDescriptiveAttributesStep : IShipStep
    {
        private readonly ShipModuleText _text;

        public RemoveDescriptiveAttributesStep(ShipModuleText text)
        {
            _text = text;
        }

        public string Name => ShipSteps.DescriptiveAttributes;

        public ShipStepReport Run(ShipWorkingSet set)
        {
            var report = new ShipStepReport { Step = Name };
            var modules = 0;
            foreach (var module in set.Modules.Where(m => !m.Removed && m.UnparsedReason == null))
            {
                var result = _text.StripDescriptiveAttributes(module.Name, set.ProjectId, module.Text, module.Type);
                if (!result.Parsed)
                {
                    module.UnparsedReason = result.Problem;
                    continue;
                }

                module.Text = result.Text;
                if (result.Removed > 0)
                {
                    modules++;
                    report.Removed += result.Removed;
                }
            }

            report.Details.Add($"{report.Removed} attribute line(s) in {modules} module(s)");
            return report;
        }
    }

    /// <summary>Takes out the document's <c>customUI/rubberduck.json</c>. The ribbon and its pictures stay.</summary>
    public class RemoveConfigStep : IShipStep
    {
        public string Name => ShipSteps.Config;

        public ShipStepReport Run(ShipWorkingSet set)
        {
            var report = new ShipStepReport { Step = Name };
            if (set.OpenPackage == null)
            {
                return report;
            }

            using (var stream = set.OpenPackage())
            {
                foreach (var resource in OfficePackage.List(stream).Where(r => r.Kind == PackageResourceKind.Config))
                {
                    set.RemovedResources.Add(resource.Name);
                    report.Removed++;
                    report.Details.Add("customUI/" + resource.Name);
                }
            }

            return report;
        }
    }
}
