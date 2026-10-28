using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Text;
using Rubberduck.PackageResources;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>
    /// What the sync does with the document's resources: the ribbon XML, the pictures and the <c>rubberduck.json</c> stored in its file.
    /// They are synced with the folder like modules are, and show up in the same list:
    /// <c>Resources\customUI14.xml</c>, <c>Resources\images\logo.png</c> and the folder's <c>rubberduck.json</c> (the one the scan found, or created).
    /// Their history is kept in the same state file, under their path as the name, which no module can have.
    /// The document side is what the <see cref="IPackageResourceService"/> shows, staged changes included, and importing stages a change like any other edit.
    /// </summary>
    public class GitSyncResources
    {
        public const string FolderName = "Resources";
        public const string ImagesFolderName = "images";

        private readonly IFileSystem _fileSystem;
        private readonly IPackageResourceService _resources;

        public GitSyncResources(IFileSystem fileSystem, IPackageResourceService resources)
        {
            _fileSystem = fileSystem;
            _resources = resources;
        }

        /// <summary>Whether the project's document has resources at all; if not, nothing here applies.</summary>
        public bool AppliesTo(string documentPath) => _resources.IsSupported(documentPath);

        private class Side
        {
            public string Name;
            public PackageResourceKind Kind;
            public string Relative;
            public byte[] Bytes;

            public string Text => Decode(Bytes);
        }

        /// <summary>Compares the document's resources with the folder's. Adds what is out of sync to the plan, and what is in sync to its baselines.</summary>
        /// <param name="configPath">The folder's rubberduck.json; its place is where the config is synced to.</param>
        public void Plan(string documentPath, string root, string configPath, GitSyncState state, GitIgnoreSet ignored, SyncPlan plan)
        {
            var document = ReadDocument(documentPath);
            var folder = ReadFolder(root, configPath);

            foreach (var key in document.Keys.Union(folder.Keys, StringComparer.OrdinalIgnoreCase).OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
            {
                document.TryGetValue(key, out var inDocument);
                folder.TryGetValue(key, out var inFolder);
                var kind = (inDocument ?? inFolder).Kind;
                var relative = inFolder?.Relative ?? inDocument.Relative;

                if (ignored.IsIgnored(relative, false))
                {
                    plan.Ignored.Add(key);
                    continue;
                }

                var item = new SyncItem
                {
                    Name = key,
                    ResourceName = (inDocument ?? inFolder).Name,
                    ResourceKind = kind,
                    RelativePath = relative,
                    ModuleExists = inDocument != null,
                    FileExists = inFolder != null,
                    ModuleSignature = inDocument == null ? null : Signature(inDocument),
                    FileSignature = inFolder == null ? null : Signature(inFolder)
                };

                state.Modules.TryGetValue(key, out var baseline);
                (item.Status, item.Suggested) = GitSyncPlanner.Classify(item, baseline);

                if (item.Status == SyncStatus.InSync)
                {
                    plan.InSyncCount++;
                    plan.InSyncBaselines[key] = BaselineOf(item);
                    continue;
                }

                //Text can be compared and merged line by line; pictures can only be taken from one side.
                if (kind != PackageResourceKind.Image && inDocument != null && inFolder != null)
                {
                    item.ModuleText = inDocument.Text;
                    item.FileText = inFolder.Text;
                }

                plan.Items.Add(item);
            }
        }

        /// <summary>Carries out the chosen action of each resource item. Stops nothing on an error: it is reported and the rest go on.</summary>
        public void Apply(string documentPath, GitSyncScan scan, IEnumerable<SyncItem> items, Action<string> backUp, GitSyncResult result)
        {
            //Ribbon XML first: pictures need one in the document.
            foreach (var item in items.OrderBy(i => i.ResourceKind == PackageResourceKind.Image ? 1 : 0).ThenBy(i => i.Name, StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    var path = PathOf(scan, item);
                    byte[] content;
                    switch (item.Chosen)
                    {
                        case SyncAction.Export:
                            content = _resources.Read(documentPath, item.ResourceName).Bytes;
                            Write(path, content, backUp);
                            result.Exported++;
                            break;

                        case SyncAction.Import:
                            content = _fileSystem.File.ReadAllBytes(path);
                            Stage(documentPath, item, content);
                            result.Imported++;
                            break;

                        case SyncAction.Merge:
                            content = Encoding.UTF8.GetBytes(item.MergedText);
                            Stage(documentPath, item, content);
                            Write(path, content, backUp);
                            result.Merged++;
                            break;

                        default:
                            continue;
                    }

                    //What each side holds now, as the next scan will see it.
                    var after = new SyncItem
                    {
                        RelativePath = item.RelativePath,
                        ModuleSignature = Signature(item.ResourceKind, _resources.Read(documentPath, item.ResourceName).Bytes),
                        FileSignature = Signature(item.ResourceKind, _fileSystem.File.ReadAllBytes(path))
                    };
                    scan.State.Modules[item.Name] = BaselineOf(after);
                }
                catch (Exception exception) when (exception is PackageResourceException || exception is System.IO.IOException || exception is UnauthorizedAccessException)
                {
                    result.Errors.Add($"Could not sync '{item.Name}': {exception.Message}");
                }
            }
        }

        private void Stage(string documentPath, SyncItem item, byte[] content)
        {
            var existing = item.ModuleExists ? _resources.Read(documentPath, item.ResourceName).Hash : null;
            switch (item.ResourceKind)
            {
                case PackageResourceKind.RibbonXml:
                    _resources.SetXml(documentPath, item.ResourceName, Decode(content), existing);
                    break;
                case PackageResourceKind.Config:
                    _resources.SetConfig(documentPath, Decode(content), existing);
                    break;
                default:
                    _resources.SetImage(documentPath, item.ResourceName, content, existing);
                    break;
            }
        }

        private void Write(string path, byte[] content, Action<string> backUp)
        {
            backUp(path);
            _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(path));
            _fileSystem.File.WriteAllBytes(path, content);
        }

        private static string PathOf(GitSyncScan scan, SyncItem item)
            => item.ResourceKind == PackageResourceKind.Config && scan.Config.Path != null
                ? scan.Config.Path
                : System.IO.Path.GetFullPath(System.IO.Path.Combine(scan.Folder, item.RelativePath.Replace('/', '\\')));

        private static ModuleBaseline BaselineOf(SyncItem item) => new ModuleBaseline
        {
            Module = item.ModuleSignature,
            File = item.FileSignature,
            Path = item.RelativePath.Replace('\\', '/'),
            Folder = string.Empty
        };

        #region Reading both sides

        private Dictionary<string, Side> ReadDocument(string documentPath)
        {
            var sides = new Dictionary<string, Side>(StringComparer.OrdinalIgnoreCase);
            foreach (var resource in _resources.List(documentPath).Where(r => r.Change != PackageResourceChange.Removed))
            {
                var key = KeyOf(resource.Kind, resource.Name);
                sides[key] = new Side
                {
                    Name = resource.Name,
                    Kind = resource.Kind,
                    Relative = RelativeOf(resource.Kind, resource.Name),
                    Bytes = _resources.Read(documentPath, resource.Name).Bytes
                };
            }

            return sides;
        }

        private Dictionary<string, Side> ReadFolder(string root, string configPath)
        {
            var sides = new Dictionary<string, Side>(StringComparer.OrdinalIgnoreCase);

            void Add(PackageResourceKind kind, string name, string path, string relative)
            {
                sides[KeyOf(kind, name)] = new Side { Name = name, Kind = kind, Relative = relative, Bytes = _fileSystem.File.ReadAllBytes(path) };
            }

            var ribbonFolder = _fileSystem.Path.Combine(root, FolderName);
            foreach (var name in PackageResourceNames.RibbonNames)
            {
                var path = _fileSystem.Path.Combine(ribbonFolder, name);
                if (_fileSystem.File.Exists(path))
                {
                    Add(PackageResourceKind.RibbonXml, name, path, RelativeOf(PackageResourceKind.RibbonXml, name));
                }
            }

            var imageFolder = _fileSystem.Path.Combine(ribbonFolder, ImagesFolderName);
            if (_fileSystem.Directory.Exists(imageFolder))
            {
                foreach (var path in _fileSystem.Directory.GetFiles(imageFolder).OrderBy(p => p, StringComparer.OrdinalIgnoreCase))
                {
                    var name = _fileSystem.Path.GetFileName(path);
                    if (PackageResourceNames.IsImageName(name))
                    {
                        Add(PackageResourceKind.Image, name, path, RelativeOf(PackageResourceKind.Image, name));
                    }
                }
            }

            var config = configPath ?? _fileSystem.Path.Combine(root, PackageResourceNames.ConfigName);
            if (_fileSystem.File.Exists(config))
            {
                Add(PackageResourceKind.Config, PackageResourceNames.ConfigName, config, RelativeToRoot(root, config));
            }

            return sides;
        }

        private string RelativeToRoot(string root, string path)
        {
            var full = _fileSystem.Path.GetFullPath(path);
            var start = _fileSystem.Path.GetFullPath(root).TrimEnd('\\') + "\\";
            return full.StartsWith(start, StringComparison.OrdinalIgnoreCase) ? full.Substring(start.Length) : full;
        }

        //The name an item has in the list and in the history; also what the user sees.
        private static string KeyOf(PackageResourceKind kind, string name)
            => kind == PackageResourceKind.Config ? PackageResourceNames.ConfigName : RelativeOf(kind, name).Replace('\\', '/');

        private static string RelativeOf(PackageResourceKind kind, string name)
            => kind == PackageResourceKind.Config ? name
                : kind == PackageResourceKind.RibbonXml ? FolderName + "\\" + name
                : FolderName + "\\" + ImagesFolderName + "\\" + name;

        #endregion

        #region Signatures

        private static string Signature(Side side) => Signature(side.Kind, side.Bytes);

        //Pictures are compared as they are. Text ignores the byte order mark, the kind of line ending and trailing blank lines, which tools add and remove freely.
        private static string Signature(PackageResourceKind kind, byte[] bytes)
            => PackageResource.HashOf(kind == PackageResourceKind.Image ? bytes : Encoding.UTF8.GetBytes(Normalized(Decode(bytes))));

        private static string Normalized(string text) => text.Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd('\n');

        private static string Decode(byte[] bytes)
        {
            using (var reader = new System.IO.StreamReader(new System.IO.MemoryStream(bytes), Encoding.UTF8, true))
            {
                return reader.ReadToEnd();
            }
        }

        #endregion
    }
}
