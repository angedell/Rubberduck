using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    public enum SyncStatus
    {
        /// <summary>The project has the module, the folder has no file for it.</summary>
        ModuleOnly,
        /// <summary>The folder has the file, the project has no such module.</summary>
        FileOnly,
        /// <summary>Only the module changed since the last sync.</summary>
        ModuleChanged,
        /// <summary>Only the file changed since the last sync.</summary>
        FileChanged,
        /// <summary>Both changed since the last sync.</summary>
        BothChanged,
        /// <summary>They differ and there is no history to tell which side is newer.</summary>
        Differs,
        InSync
    }

    public enum SyncAction
    {
        None,
        /// <summary>Write the module to the file.</summary>
        Export,
        /// <summary>Load the file into the project.</summary>
        Import,
        /// <summary>Take each difference from the side the user picked, and write the result to both.</summary>
        Merge
    }

    /// <summary>A module of the project, as the VBE exports it.</summary>
    public class ModuleInput
    {
        public string Name { get; set; }

        /// <summary>The extension its file has, e.g. ".cls".</summary>
        public string Extension { get; set; }

        public string ExportedText { get; set; }
    }

    /// <summary>A source file found in the sync folder.</summary>
    public class FileInput
    {
        public string Name { get; set; }

        /// <summary>The path relative to the sync folder, with either kind of separator.</summary>
        public string RelativePath { get; set; }

        public string Text { get; set; }
    }

    public class SyncItem
    {
        public string Name { get; set; }
        public SyncStatus Status { get; set; }

        /// <summary>What the history says to do; None if the user has to decide.</summary>
        public SyncAction Suggested { get; set; }

        /// <summary>The path of the file relative to the sync folder, or where it would go if the module was exported.</summary>
        public string RelativePath { get; set; }

        public bool ModuleExists { get; set; }
        public bool FileExists { get; set; }

        /// <summary>Exporting would put the file somewhere else than where it is now, because the module's @Folder changed.</summary>
        public bool MovesFile { get; set; }

        public string ModuleSignature { get; set; }
        public string FileSignature { get; set; }

        public List<SyncViolation> Violations { get; } = new List<SyncViolation>();

        /// <summary>The module as the VBE exports it. Kept for the items that are out of sync and exist on both sides, so they can be compared line by line.</summary>
        public string ModuleText { get; set; }

        /// <summary>The file, ready for import. Kept like <see cref="ModuleText"/>.</summary>
        public string FileText { get; set; }

        /// <summary>The text the user merged from both sides, for <see cref="SyncAction.Merge"/>.</summary>
        public string MergedText { get; set; }

        /// <summary>What the merged text breaks.</summary>
        public List<SyncViolation> MergedViolations { get; } = new List<SyncViolation>();

        public bool CanCompare => ModuleText != null && FileText != null;

        /// <summary>For a resource of the document (ribbon XML, picture, config) rather than a module: its name in the package. Null for modules.</summary>
        public string ResourceName { get; set; }

        public Rubberduck.PackageResources.PackageResourceKind ResourceKind { get; set; }

        public bool IsResource => ResourceName != null;

        /// <summary>The action the user went for, set by the review.</summary>
        public SyncAction Chosen { get; set; }

        /// <summary>Whether using the given side as the source is stopped by a rule the config enforces.</summary>
        public bool IsBlockedFor(SyncAction action)
        {
            if (action == SyncAction.Merge)
            {
                return MergedViolations.Any(violation => violation.Blocking);
            }

            var side = action == SyncAction.Export ? SyncSide.Module : SyncSide.File;
            return action != SyncAction.None && Violations.Any(violation => violation.Blocking && violation.Side == side);
        }

        public bool IsBlocked => Chosen != SyncAction.None && IsBlockedFor(Chosen);
    }

    public class SyncPlan
    {
        /// <summary>The modules and files that are not in sync.</summary>
        public List<SyncItem> Items { get; } = new List<SyncItem>();

        public int InSyncCount { get; set; }

        /// <summary>The history to record for the modules that are in sync, so that the next change can be traced to a side.</summary>
        public Dictionary<string, ModuleBaseline> InSyncBaselines { get; } = new Dictionary<string, ModuleBaseline>(StringComparer.OrdinalIgnoreCase);

        /// <summary>What the ignore list left out.</summary>
        public List<string> Ignored { get; } = new List<string>();
    }

    /// <summary>
    /// Works out what is out of sync between the project and the folder. Pure: it is handed the text of both sides and touches nothing,
    /// so the modules can be compared on several threads at once.
    /// </summary>
    public class GitSyncPlanner
    {
        private readonly IIndenterSettings _settings;
        private readonly ISimpleIndenter _indenter;
        private readonly SyncConfig _config;
        private readonly SignatureCache _cache;
        private int _completed;

        public GitSyncPlanner(IIndenterSettings settings, ISimpleIndenter indenter, SyncConfig config, SignatureCache cache = null)
        {
            _settings = settings;
            _indenter = indenter;
            _config = config ?? new SyncConfig();
            _cache = cache ?? new SignatureCache(settings, null, _config.IgnoreCaseOutsideStrings);
        }

        /// <summary>How many modules and files have been compared so far. Safe to read while <see cref="Plan"/> runs.</summary>
        public int Completed => _completed;

        public SyncPlan Plan(IEnumerable<ModuleInput> modules, IEnumerable<FileInput> files, GitSyncState state, GitIgnoreSet ignored, CancellationToken token = default(CancellationToken))
        {
            var fileByName = new Dictionary<string, FileInput>(StringComparer.OrdinalIgnoreCase);
            foreach (var file in files.Where(f => f.Name != null))
            {
                if (!fileByName.ContainsKey(file.Name))
                {
                    fileByName.Add(file.Name, file);
                }
            }

            var moduleByName = modules.ToDictionary(m => m.Name, StringComparer.OrdinalIgnoreCase);
            var names = moduleByName.Keys.Union(fileByName.Keys, StringComparer.OrdinalIgnoreCase).OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();
            Total = names.Count;
            _completed = 0;

            //Each name is independent, so they are examined in parallel. The state and the ignore rules are only read.
            var results = names.AsParallel()
                .AsOrdered()
                .WithCancellation(token)
                .Select(name =>
                {
                    moduleByName.TryGetValue(name, out var module);
                    fileByName.TryGetValue(name, out var file);
                    state.Modules.TryGetValue(name, out var baseline);
                    var examined = Examine(name, module, file, baseline, ignored);
                    Interlocked.Increment(ref _completed);
                    return new Examined { Name = name, Item = examined.Item, Module = module, Ignored = examined.Ignored };
                })
                .ToList();

            var plan = new SyncPlan();
            foreach (var result in results)
            {
                if (result.Ignored)
                {
                    plan.Ignored.Add(result.Name);
                }
                else if (result.Item.Status == SyncStatus.InSync)
                {
                    plan.InSyncCount++;
                    plan.InSyncBaselines[result.Name] = BaselineOf(result.Item, result.Module);
                }
                else
                {
                    plan.Items.Add(result.Item);
                }
            }

            return plan;
        }

        /// <summary>The number of modules and files <see cref="Plan"/> goes through.</summary>
        public int Total { get; private set; }

        private class Examined
        {
            public string Name;
            public SyncItem Item;
            public ModuleInput Module;
            public bool Ignored;
        }

        private (SyncItem Item, bool Ignored) Examine(string name, ModuleInput module, FileInput file, ModuleBaseline baseline, GitIgnoreSet ignored)
        {
            var item = new SyncItem
            {
                Name = name,
                ModuleExists = module != null,
                FileExists = file != null
            };

            //Where the file is, or would go, is known without reading either side; ignored items cost nothing more.
            string folder = null;
            if (module != null)
            {
                folder = GitSyncText.GetFolderAnnotation(SyncText.CodeOf(module.ExportedText)) ?? string.Empty;
                var target = TargetPath(name, module.Extension, folder, baseline);
                item.RelativePath = file == null ? target : file.RelativePath;
                item.MovesFile = file != null && !string.Equals(Normalize(target), Normalize(file.RelativePath), StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                item.RelativePath = file.RelativePath;
            }

            if (ignored.IsIgnored(item.RelativePath, false))
            {
                return (item, true);
            }

            string fileText = null;
            var directory = string.Empty;
            if (file != null)
            {
                directory = System.IO.Path.GetDirectoryName(file.RelativePath.Replace('/', '\\')) ?? string.Empty;
                fileText = SyncText.PrepareForImport(file.Text, directory);
                item.FileSignature = _cache.Signature(fileText, _settings, _indenter);
            }

            if (module != null)
            {
                item.ModuleSignature = _cache.Signature(module.ExportedText, _settings, _indenter);
            }

            (item.Status, item.Suggested) = Classify(item, baseline);

            //The rules only matter for what is going to be listed.
            if (item.Status != SyncStatus.InSync)
            {
                if (file != null && module != null)
                {
                    item.ModuleText = module.ExportedText;
                    item.FileText = fileText;
                }

                if (file != null)
                {
                    item.Violations.AddRange(SyncHygiene.Check(name, fileText, SyncSide.File, _config, directory.Length > 0));
                }

                if (module != null)
                {
                    item.Violations.AddRange(SyncHygiene.Check(name, module.ExportedText, SyncSide.Module, _config));
                }
            }

            return (item, false);
        }

        public static ModuleBaseline BaselineOf(SyncItem item, ModuleInput module)
        {
            return new ModuleBaseline
            {
                Module = item.ModuleSignature,
                File = item.FileSignature,
                Path = item.RelativePath.Replace('\\', '/'),
                Folder = module == null ? string.Empty : GitSyncText.GetFolderAnnotation(SyncText.CodeOf(module.ExportedText)) ?? string.Empty
            };
        }

        //The file stays where it is, unless the module was moved to another virtual folder.
        public static string TargetPath(string name, string extension, string folder, ModuleBaseline baseline)
        {
            if (baseline?.Path != null && baseline.Folder == folder)
            {
                return baseline.Path;
            }

            var directory = GitSyncText.RelativeDirectory(folder);
            return directory.Length == 0 ? name + extension : directory + "\\" + name + extension;
        }

        private static string Normalize(string path) => path.Replace('/', '\\');

        /// <summary>Which side changed since the last sync, from the signatures of both sides and the history.</summary>
        public static (SyncStatus, SyncAction) Classify(SyncItem item, ModuleBaseline baseline)
        {
            if (!item.FileExists)
            {
                return (SyncStatus.ModuleOnly, SyncAction.Export);
            }

            if (!item.ModuleExists)
            {
                return (SyncStatus.FileOnly, SyncAction.Import);
            }

            if (item.ModuleSignature == item.FileSignature && !item.MovesFile)
            {
                return (SyncStatus.InSync, SyncAction.None);
            }

            if (baseline == null)
            {
                return (SyncStatus.Differs, SyncAction.None);
            }

            var moduleChanged = item.ModuleSignature != baseline.Module;
            var fileChanged = item.FileSignature != baseline.File;

            if (moduleChanged && !fileChanged)
            {
                return (SyncStatus.ModuleChanged, SyncAction.Export);
            }

            if (fileChanged && !moduleChanged)
            {
                return (SyncStatus.FileChanged, SyncAction.Import);
            }

            if (moduleChanged)
            {
                return item.ModuleSignature == item.FileSignature
                    ? (SyncStatus.ModuleChanged, SyncAction.Export) //Both ended up the same; only the file's place is left to fix.
                    : (SyncStatus.BothChanged, SyncAction.None);
            }

            //Neither changed since the last sync, and yet they differ: the history cannot be trusted, e.g. the VBE rewrote the header.
            return item.MovesFile && item.ModuleSignature == item.FileSignature
                ? (SyncStatus.ModuleChanged, SyncAction.Export)
                : (SyncStatus.Differs, SyncAction.None);
        }
    }
}
