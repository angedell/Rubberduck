using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;
using NLog;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using FileAccess = System.IO.FileAccess;
using FileMode = System.IO.FileMode;
using FileShare = System.IO.FileShare;
using IOException = System.IO.IOException;
using MemoryStream = System.IO.MemoryStream;
using Stream = System.IO.Stream;
using StreamReader = System.IO.StreamReader;

namespace Rubberduck.PackageResources
{
    public enum PackageApplyOutcome
    {
        /// <summary>Nothing was waiting.</summary>
        NothingPending,

        /// <summary>The changes are in the file now.</summary>
        Applied,

        /// <summary>The file is still open in its application; try again after it is closed.</summary>
        Busy,

        /// <summary>The file cannot be changed (it is gone, or not a package any more). The changes are kept.</summary>
        Failed
    }

    /// <summary>
    /// The ribbon XML and pictures that live inside a document's file, as the Code Explorer and the API see them.
    /// The application keeps the file locked while the document is open, so changes are not written to it: they are staged, and the listing and reads
    /// already include them. They go into the file when it is free, i.e. after the document is closed.
    /// </summary>
    public interface IPackageResourceService
    {
        /// <summary>Raised when staged changes were made, discarded or applied. May come from any thread.</summary>
        event EventHandler Changed;

        /// <summary>Whether the file at that path is a document with a ribbon folder (a saved, Open XML file).</summary>
        bool IsSupported(string documentPath);

        /// <summary>The ribbon XML and pictures, staged changes included. Ribbon XML first, then pictures by name.</summary>
        IReadOnlyList<PackageResource> List(string documentPath);

        PackageResourceContent Read(string documentPath, string name);

        /// <param name="expectedHash">The hash from the read the edit is based on. Not needed when the part does not exist yet.</param>
        PackageResource SetXml(string documentPath, string name, string xml, string expectedHash);

        /// <summary>Stages the document's <c>rubberduck.json</c>. The JSON has to be well-formed; settings it gets wrong are reported by <see cref="CheckConfig"/>, not refused.</summary>
        /// <param name="expectedHash">The hash from the read the edit is based on. Not needed when the document has no config yet.</param>
        PackageResource SetConfig(string documentPath, string json, string expectedHash);

        /// <summary>
        /// Gives the document a <c>rubberduck.json</c> if it has none, staged like any other change. Only the first call per document in this session does anything,
        /// so discarding the new config does not bring it back until Rubberduck is restarted.
        /// </summary>
        /// <returns>True if one was created now.</returns>
        bool EnsureConfig(string documentPath);

        /// <summary>What is wrong with the config's settings (unknown keys, bad values), without refusing it. Empty if nothing is.</summary>
        IReadOnlyList<string> CheckConfig(string json);

        /// <param name="expectedHash">Needed to replace a picture that exists; not to add one.</param>
        PackageResource SetImage(string documentPath, string name, byte[] bytes, string expectedHash);

        /// <summary>Stages the removal of pictures. A picture that was only staged is simply dropped.</summary>
        void RemoveImages(string documentPath, IEnumerable<string> names);

        bool HasPending(string documentPath);

        /// <summary>The documents that have staged changes, whether or not they are open.</summary>
        IReadOnlyList<string> PendingDocuments();

        /// <summary>Throws the staged changes of the document away.</summary>
        void Discard(string documentPath);

        /// <summary>Writes the staged changes into the file, if it is free.</summary>
        PackageApplyOutcome TryApply(string documentPath, out string message);
    }

    public class PackageResourceService : IPackageResourceService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IFileSystem _fileSystem;
        private readonly string _journalFolder;
        private readonly string _backupFolder;
        private readonly object _lock = new object();
        private readonly Func<string> _defaultConfig;
        private readonly HashSet<string> _configEnsured = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public PackageResourceService(IFileSystem fileSystem, Func<string> defaultConfig = null)
            : this(fileSystem,
                fileSystem.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Rubberduck", "PendingResources"),
                fileSystem.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Rubberduck", "ResourceBackups"),
                defaultConfig)
        {
        }

        /// <param name="defaultConfig">The text of a new rubberduck.json, e.g. with the user's indenter settings. A plain default if null.</param>
        public PackageResourceService(IFileSystem fileSystem, string journalFolder, string backupFolder, Func<string> defaultConfig = null)
        {
            _fileSystem = fileSystem;
            _journalFolder = journalFolder;
            _backupFolder = backupFolder;
            _defaultConfig = defaultConfig ?? (() => ProjectConfigTemplate.Create(new IndenterSettings(true)));
        }

        public event EventHandler Changed;

        #region Staged changes

        private class StagedItem
        {
            public PackageResourceKind Kind { get; set; }

            /// <summary>True: take it out of the file. False: write <see cref="Content"/> into it.</summary>
            public bool Remove { get; set; }

            public byte[] Content { get; set; }
        }

        private class Journal
        {
            public string Document { get; set; }

            public Dictionary<string, StagedItem> Items { get; set; } = new Dictionary<string, StagedItem>(StringComparer.OrdinalIgnoreCase);
        }

        private string JournalFile(string documentPath)
        {
            using (var sha = SHA1.Create())
            {
                var key = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(Normalize(documentPath).ToLowerInvariant())), 0, 8).Replace("-", string.Empty);
                return _fileSystem.Path.Combine(_journalFolder, key + ".json");
            }
        }

        private string Normalize(string documentPath)
        {
            try
            {
                return _fileSystem.Path.GetFullPath(documentPath);
            }
            catch (Exception)
            {
                return documentPath ?? string.Empty;
            }
        }

        private Journal LoadJournal(string documentPath)
        {
            var file = JournalFile(documentPath);
            try
            {
                if (_fileSystem.File.Exists(file))
                {
                    var journal = JsonConvert.DeserializeObject<Journal>(_fileSystem.File.ReadAllText(file));
                    if (journal?.Items != null)
                    {
                        journal.Items = new Dictionary<string, StagedItem>(journal.Items, StringComparer.OrdinalIgnoreCase);
                        return journal;
                    }
                }
            }
            catch (Exception exception)
            {
                Logger.Warn(exception, "Could not read the staged resource changes in {0}.", file);
            }

            return new Journal { Document = Normalize(documentPath) };
        }

        private void SaveJournal(string documentPath, Journal journal)
        {
            var file = JournalFile(documentPath);
            if (journal.Items.Count == 0)
            {
                DeleteJournal(documentPath);
                return;
            }

            try
            {
                _fileSystem.Directory.CreateDirectory(_journalFolder);
                _fileSystem.File.WriteAllText(file, JsonConvert.SerializeObject(journal));
            }
            catch (Exception exception)
            {
                throw new PackageResourceException(PackageResourceError.Unavailable, $"Could not keep the change for later: {exception.Message}", exception);
            }
        }

        private void DeleteJournal(string documentPath)
        {
            var file = JournalFile(documentPath);
            try
            {
                if (_fileSystem.File.Exists(file))
                {
                    _fileSystem.File.Delete(file);
                }
            }
            catch (Exception exception)
            {
                Logger.Warn(exception, "Could not delete the staged resource changes in {0}.", file);
            }
        }

        private void OnChanged() => Changed?.Invoke(this, EventArgs.Empty);

        #endregion

        public bool IsSupported(string documentPath)
        {
            if (string.IsNullOrWhiteSpace(documentPath))
            {
                return false;
            }

            try
            {
                return PackageResourceNames.DocumentExtensions.Contains(_fileSystem.Path.GetExtension(documentPath), StringComparer.OrdinalIgnoreCase)
                       && _fileSystem.File.Exists(documentPath);
            }
            catch (ArgumentException)
            {
                return false;
            }
        }

        public bool HasPending(string documentPath)
        {
            lock (_lock)
            {
                return LoadJournal(documentPath).Items.Count > 0;
            }
        }

        public IReadOnlyList<string> PendingDocuments()
        {
            lock (_lock)
            {
                var documents = new List<string>();
                try
                {
                    if (!_fileSystem.Directory.Exists(_journalFolder))
                    {
                        return documents;
                    }

                    foreach (var file in _fileSystem.Directory.GetFiles(_journalFolder, "*.json"))
                    {
                        var journal = JsonConvert.DeserializeObject<Journal>(_fileSystem.File.ReadAllText(file));
                        if (journal?.Items != null && journal.Items.Count > 0 && !string.IsNullOrEmpty(journal.Document))
                        {
                            documents.Add(journal.Document);
                        }
                    }
                }
                catch (Exception exception)
                {
                    Logger.Warn(exception, "Could not list the staged resource changes.");
                }

                return documents;
            }
        }

        public IReadOnlyList<PackageResource> List(string documentPath)
        {
            lock (_lock)
            {
                return Merge(committed: Committed(documentPath), staged: LoadJournal(documentPath));
            }
        }

        private static List<PackageResource> Merge(List<PackageResource> committed, Journal staged)
        {
            var merged = new List<PackageResource>();
            foreach (var resource in committed)
            {
                if (!staged.Items.TryGetValue(resource.Name, out var item))
                {
                    merged.Add(resource);
                }
                else if (item.Remove)
                {
                    resource.Change = PackageResourceChange.Removed;
                    merged.Add(resource);
                }
                else
                {
                    merged.Add(Describe(resource.Name, item.Kind, item.Content, PackageResourceChange.Modified));
                }
            }

            foreach (var item in staged.Items.Where(i => !i.Value.Remove && committed.All(c => !c.Name.Equals(i.Key, StringComparison.OrdinalIgnoreCase))))
            {
                merged.Add(Describe(item.Key, item.Value.Kind, item.Value.Content, PackageResourceChange.Added));
            }

            return merged
                .OrderBy(r => r.Kind)
                .ThenBy(r => r.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static PackageResource Describe(string name, PackageResourceKind kind, byte[] content, PackageResourceChange change)
            => new PackageResource { Name = name, Kind = kind, Size = content.Length, Hash = PackageResource.HashOf(content), Change = change };

        public PackageResourceContent Read(string documentPath, string name)
        {
            lock (_lock)
            {
                return ReadLocked(documentPath, name);
            }
        }

        private PackageResourceContent ReadLocked(string documentPath, string name)
        {
            var staged = LoadJournal(documentPath);
            var kind = PackageResourceNames.KindOf(name);
            var change = PackageResourceChange.None;
            byte[] bytes = null;

            if (staged.Items.TryGetValue(name, out var item) && !item.Remove)
            {
                bytes = item.Content;
                change = PackageResourceChange.Added;
            }

            var inFile = Committed(documentPath).FirstOrDefault(r => r.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (inFile != null)
            {
                name = inFile.Name;
                if (bytes != null)
                {
                    change = PackageResourceChange.Modified;
                }
                else
                {
                    change = item != null ? PackageResourceChange.Removed : PackageResourceChange.None;
                    bytes = FromFile(documentPath, name);
                }
            }

            if (bytes == null)
            {
                throw new PackageResourceException(PackageResourceError.NotFound, $"The document has no resource '{name}'.");
            }

            return new PackageResourceContent
            {
                Name = name,
                Kind = kind,
                Size = bytes.Length,
                Hash = PackageResource.HashOf(bytes),
                Change = change,
                Bytes = bytes,
                Text = kind != PackageResourceKind.Image ? Decode(bytes) : null
            };
        }

        private static string Decode(byte[] bytes)
        {
            using (var reader = new StreamReader(new MemoryStream(bytes), Encoding.UTF8, true))
            {
                return reader.ReadToEnd();
            }
        }

        #region Staging

        public PackageResource SetXml(string documentPath, string name, string xml, string expectedHash)
        {
            if (!PackageResourceNames.IsRibbonName(name))
            {
                throw new PackageResourceException(PackageResourceError.BadRequest,
                    $"Ribbon XML is called {PackageResourceNames.Ribbon2007} (Office 2007) or {PackageResourceNames.Ribbon2010} (Office 2010 and later), not '{name}'.");
            }

            name = PackageResourceNames.RibbonNames.First(n => n.Equals(name, StringComparison.OrdinalIgnoreCase));
            var bytes = Validated(name, xml);

            lock (_lock)
            {
                var resource = Stage(documentPath, name, PackageResourceKind.RibbonXml, bytes, expectedHash);
                OnChanged();
                return resource;
            }
        }

        public PackageResource SetConfig(string documentPath, string json, string expectedHash)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                throw new PackageResourceException(PackageResourceError.BadRequest, "The config is empty.");
            }

            try
            {
                if (!(Newtonsoft.Json.Linq.JToken.Parse(json) is Newtonsoft.Json.Linq.JObject))
                {
                    throw new PackageResourceException(PackageResourceError.BadRequest, "The config must be a JSON object.");
                }
            }
            catch (Newtonsoft.Json.JsonException exception)
            {
                throw new PackageResourceException(PackageResourceError.BadRequest, $"The config is not valid JSON: {exception.Message}", exception);
            }

            lock (_lock)
            {
                var resource = Stage(documentPath, PackageResourceNames.ConfigName, PackageResourceKind.Config, new UTF8Encoding(false).GetBytes(json), expectedHash);
                OnChanged();
                return resource;
            }
        }

        public bool EnsureConfig(string documentPath)
        {
            lock (_lock)
            {
                if (!IsSupported(documentPath) || !_configEnsured.Add(Normalize(documentPath)))
                {
                    return false;
                }

                try
                {
                    if (List(documentPath).Any(r => r.Kind == PackageResourceKind.Config && r.Change != PackageResourceChange.Removed))
                    {
                        return false;
                    }

                    var journal = LoadJournal(documentPath);
                    journal.Document = Normalize(documentPath);
                    journal.Items[PackageResourceNames.ConfigName] = new StagedItem { Kind = PackageResourceKind.Config, Content = new UTF8Encoding(false).GetBytes(_defaultConfig()) };
                    SaveJournal(documentPath, journal);
                }
                catch (PackageResourceException exception)
                {
                    //Locked for now; the next session tries again.
                    Logger.Debug(exception, "No rubberduck.json created for {0}.", documentPath);
                    return false;
                }
            }

            OnChanged();
            return true;
        }

        public IReadOnlyList<string> CheckConfig(string json)
            => new ProjectConfigReader(_fileSystem).Parse(json, null).Errors.ToList();

        public PackageResource SetImage(string documentPath, string name, byte[] bytes, string expectedHash)
        {
            if (!PackageResourceNames.IsImageName(name))
            {
                throw new PackageResourceException(PackageResourceError.BadRequest,
                    $"'{name}' is not a usable picture name. Start with a letter, use letters, digits, '_', '-' and '.', and end in {string.Join(", ", PackageResourceNames.ImageExtensions)}.");
            }

            if (bytes == null || bytes.Length == 0)
            {
                throw new PackageResourceException(PackageResourceError.BadRequest, "The picture is empty.");
            }

            if (bytes.Length > PackageResourceNames.MaxImageBytes)
            {
                throw new PackageResourceException(PackageResourceError.BadRequest, $"The picture is larger than {PackageResourceNames.MaxImageBytes / (1024 * 1024)} MB.");
            }

            if (!LooksLikeImage(_fileSystem.Path.GetExtension(name), bytes))
            {
                throw new PackageResourceException(PackageResourceError.BadRequest, $"The content is not a {_fileSystem.Path.GetExtension(name)} picture.");
            }

            lock (_lock)
            {
                var staged = LoadJournal(documentPath);
                var listing = Merge(Committed(documentPath), staged);
                if (!listing.Any(r => r.Kind == PackageResourceKind.RibbonXml && r.Change != PackageResourceChange.Removed))
                {
                    throw new PackageResourceException(PackageResourceError.Conflict,
                        "The document has no ribbon XML yet, and a picture is only useful to it. Add the ribbon XML first.");
                }

                var resource = Stage(documentPath, name, PackageResourceKind.Image, bytes, expectedHash);
                OnChanged();
                return resource;
            }
        }

        //Called with the lock held.
        private PackageResource Stage(string documentPath, string name, PackageResourceKind kind, byte[] bytes, string expectedHash)
        {
            var committed = Committed(documentPath);
            var journal = LoadJournal(documentPath);

            var current = Merge(committed, journal).FirstOrDefault(r => r.Name.Equals(name, StringComparison.OrdinalIgnoreCase) && r.Change != PackageResourceChange.Removed);
            if (current != null)
            {
                if (string.IsNullOrWhiteSpace(expectedHash))
                {
                    throw new PackageResourceException(PackageResourceError.HashRequired,
                        $"'{name}' exists. Send the hash from your last read of it, so changes made in the meantime are not overwritten.");
                }

                var hash = current.Hash ?? PackageResource.HashOf(FromFile(documentPath, current.Name));
                if (!string.Equals(hash, expectedHash, StringComparison.OrdinalIgnoreCase))
                {
                    throw new PackageResourceException(PackageResourceError.Stale, $"'{name}' changed since you read it. Read it again.");
                }
            }

            //Naming a committed resource with other casing must replace it, not add a second one.
            var existingName = committed.FirstOrDefault(c => c.Name.Equals(name, StringComparison.OrdinalIgnoreCase))?.Name ?? name;
            journal.Document = Normalize(documentPath);
            journal.Items[existingName] = new StagedItem { Kind = kind, Content = bytes };
            SaveJournal(documentPath, journal);

            return Describe(existingName, kind, bytes, committed.Any(c => c.Name.Equals(existingName, StringComparison.OrdinalIgnoreCase)) ? PackageResourceChange.Modified : PackageResourceChange.Added);
        }

        public void RemoveImages(string documentPath, IEnumerable<string> names)
        {
            var requested = names.ToList();
            lock (_lock)
            {
                var committed = Committed(documentPath);
                var journal = LoadJournal(documentPath);
                var listing = Merge(committed, journal);

                foreach (var name in requested)
                {
                    var found = listing.FirstOrDefault(r => r.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                    if (found == null || found.Change == PackageResourceChange.Removed)
                    {
                        throw new PackageResourceException(PackageResourceError.NotFound, $"The document has no picture '{name}'.");
                    }

                    if (found.Kind != PackageResourceKind.Image)
                    {
                        throw new PackageResourceException(PackageResourceError.BadRequest, "Ribbon XML and rubberduck.json cannot be removed, only changed.");
                    }
                }

                foreach (var name in requested)
                {
                    var inFile = committed.FirstOrDefault(c => c.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                    if (inFile == null)
                    {
                        journal.Items.Remove(name);
                    }
                    else
                    {
                        journal.Items[inFile.Name] = new StagedItem { Kind = PackageResourceKind.Image, Remove = true };
                    }
                }

                journal.Document = Normalize(documentPath);
                SaveJournal(documentPath, journal);
                OnChanged();
            }
        }

        public void Discard(string documentPath)
        {
            lock (_lock)
            {
                DeleteJournal(documentPath);
            }

            OnChanged();
        }

        private static byte[] Validated(string name, string xml)
        {
            var problem = RibbonXml.Check(name, xml);
            if (problem != null)
            {
                throw new PackageResourceException(PackageResourceError.BadRequest,
                    problem.Line > 0 ? $"Line {problem.Line}, column {problem.Column}: {problem.Message}" : problem.Message);
            }

            return RibbonXml.ToBytes(xml);
        }

        private static bool LooksLikeImage(string extension, byte[] bytes)
        {
            bool Starts(params byte[] signature) => bytes.Length >= signature.Length && signature.SequenceEqual(bytes.Take(signature.Length));

            switch (extension.ToLowerInvariant())
            {
                case ".png": return Starts(0x89, 0x50, 0x4E, 0x47);
                case ".jpg":
                case ".jpeg": return Starts(0xFF, 0xD8);
                case ".gif": return Starts(0x47, 0x49, 0x46, 0x38);
                case ".bmp": return Starts(0x42, 0x4D);
                case ".ico": return Starts(0x00, 0x00, 0x01, 0x00);
                default: return false;
            }
        }

        #endregion

        #region The file

        private List<PackageResource> Committed(string documentPath)
        {
            if (!IsSupported(documentPath))
            {
                throw new PackageResourceException(PackageResourceError.Unavailable, "The document has not been saved as an Open XML file (xlsm, docm, pptm...), so it has no resources to show.");
            }

            try
            {
                using (var stream = OpenForReading(documentPath))
                {
                    return OfficePackage.List(stream);
                }
            }
            catch (IOException exception)
            {
                throw new PackageResourceException(PackageResourceError.Unavailable, $"The document's file cannot be read right now: {exception.Message}", exception);
            }
        }

        private byte[] FromFile(string documentPath, string name)
        {
            try
            {
                using (var stream = OpenForReading(documentPath))
                {
                    return OfficePackage.Read(stream, name);
                }
            }
            catch (IOException exception)
            {
                throw new PackageResourceException(PackageResourceError.Unavailable, $"The document's file cannot be read right now: {exception.Message}", exception);
            }
        }

        //The application has the file open for reading and shares reads with others, so that is all we ask for.
        private Stream OpenForReading(string documentPath)
            => _fileSystem.File.Open(documentPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);

        public PackageApplyOutcome TryApply(string documentPath, out string message)
        {
            lock (_lock)
            {
                message = null;
                var journal = LoadJournal(documentPath);
                if (journal.Items.Count == 0)
                {
                    return PackageApplyOutcome.NothingPending;
                }

                if (!_fileSystem.File.Exists(documentPath))
                {
                    message = "The document's file no longer exists.";
                    return PackageApplyOutcome.Failed;
                }

                var temporary = documentPath + ".rd-resources.tmp";
                try
                {
                    //Asking for exclusive access fails while the application has the document open, before anything is copied or backed up.
                    using (_fileSystem.File.Open(documentPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    {
                    }

                    var edits = journal.Items.ToDictionary(i => i.Key, i => i.Value.Remove ? null : i.Value.Content, StringComparer.OrdinalIgnoreCase);

                    using (var source = OpenForReading(documentPath))
                    using (var target = _fileSystem.File.Open(temporary, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
                    {
                        OfficePackage.Patch(source, target, edits);
                    }

                    var backup = Backup(documentPath);

                    //Fails with a sharing violation while the application still has the document open; nothing has changed then.
                    _fileSystem.File.Replace(temporary, documentPath, null);
                    Logger.Info("Applied {0} resource change(s) to {1} (original kept in {2}).", edits.Count, documentPath, backup);
                }
                catch (UnauthorizedAccessException)
                {
                    return Busy(temporary, out message);
                }
                catch (IOException exception) when (IsSharing(exception))
                {
                    return Busy(temporary, out message);
                }
                catch (Exception exception)
                {
                    Logger.Warn(exception, "Could not apply the staged resource changes to {0}.", documentPath);
                    TryDelete(temporary);
                    message = exception.Message;
                    return PackageApplyOutcome.Failed;
                }

                DeleteJournal(documentPath);
            }

            OnChanged();
            return PackageApplyOutcome.Applied;
        }

        private PackageApplyOutcome Busy(string temporary, out string message)
        {
            TryDelete(temporary);
            message = "The document is still open.";
            return PackageApplyOutcome.Busy;
        }

        private static bool IsSharing(IOException exception)
        {
            //ERROR_SHARING_VIOLATION and ERROR_LOCK_VIOLATION
            var code = System.Runtime.InteropServices.Marshal.GetHRForException(exception) & 0xFFFF;
            return code == 32 || code == 33;
        }

        private string Backup(string documentPath)
        {
            var folder = _fileSystem.Path.Combine(_backupFolder, DateTime.Now.ToString("yyyyMMdd-HHmmss"));
            _fileSystem.Directory.CreateDirectory(folder);
            var copy = _fileSystem.Path.Combine(folder, _fileSystem.Path.GetFileName(documentPath));
            using (var source = OpenForReading(documentPath))
            using (var target = _fileSystem.File.Create(copy))
            {
                source.CopyTo(target);
            }

            return copy;
        }

        private void TryDelete(string path)
        {
            try
            {
                if (_fileSystem.File.Exists(path))
                {
                    _fileSystem.File.Delete(path);
                }
            }
            catch (Exception exception)
            {
                Logger.Debug(exception, "Could not delete {0}.", path);
            }
        }

        #endregion
    }
}
