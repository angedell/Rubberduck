using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using Newtonsoft.Json;
using NLog;
using Rubberduck.SmartIndenter;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>
    /// Remembers the signature of text that was seen before, so that indenting it is not done again. The key is a hash of the text itself and
    /// of the indenter settings, so a hit is always right: nothing is guessed from file times, and a change to the text or the settings is simply a miss.
    /// It is kept per user, outside the repository, in <c>%LocalAppData%\Rubberduck\gitsync-cache.json</c>.
    /// Safe to use from several threads.
    /// </summary>
    public class SignatureCache
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        //Changing how signatures are computed makes every older entry wrong; bump this to drop them.
        private const string AlgorithmVersion = "1";
        private const int MaximumEntries = 20000;

        private static readonly PropertyInfo[] SettingProperties = typeof(IIndenterSettings)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead)
            .OrderBy(p => p.Name, StringComparer.Ordinal)
            .ToArray();

        private readonly ConcurrentDictionary<string, string> _entries;
        private readonly ConcurrentDictionary<string, bool> _used = new ConcurrentDictionary<string, bool>();
        private readonly string _settingsKey;
        private readonly bool _ignoreCase;
        private int _hits;
        private int _misses;

        public SignatureCache(IIndenterSettings settings, IDictionary<string, string> entries = null, bool ignoreCaseOutsideStrings = false)
        {
            _ignoreCase = ignoreCaseOutsideStrings;
            _settingsKey = AlgorithmVersion + "|case=" + (ignoreCaseOutsideStrings ? "fold" : "exact") + "|" + KeyOf(settings);
            _entries = new ConcurrentDictionary<string, string>(entries ?? new Dictionary<string, string>());
        }

        public int Hits => _hits;
        public int Misses => _misses;

        /// <summary>The settings as one string; two settings with the same string indent the same way.</summary>
        public static string KeyOf(IIndenterSettings settings)
        {
            return settings == null
                ? string.Empty
                : string.Join(";", SettingProperties.Select(p => p.Name + "=" + p.GetValue(settings)));
        }

        /// <summary>The signature of the text, from the cache if the same text was seen with the same settings.</summary>
        public string Signature(string text, IIndenterSettings settings, ISimpleIndenter indenter)
        {
            var key = KeyFor(text);
            _used[key] = true;
            if (_entries.TryGetValue(key, out var known))
            {
                Interlocked.Increment(ref _hits);
                return known;
            }

            Interlocked.Increment(ref _misses);
            var signature = SyncText.Signature(text, settings, indenter, _ignoreCase);
            _entries[key] = signature;
            return signature;
        }

        private string KeyFor(string text)
        {
            using (var sha = SHA1.Create())
            {
                var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(_settingsKey + "\u0001" + text));
                return Convert.ToBase64String(hash, 0, 12);
            }
        }

        #region Persistence

        public static string DefaultPath => System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Rubberduck", "gitsync-cache.json");

        public static SignatureCache Load(IFileSystem fileSystem, IIndenterSettings settings, string path = null, bool ignoreCaseOutsideStrings = false)
        {
            path = path ?? DefaultPath;
            try
            {
                if (fileSystem.File.Exists(path))
                {
                    var entries = JsonConvert.DeserializeObject<Dictionary<string, string>>(fileSystem.File.ReadAllText(path));
                    return new SignatureCache(settings, entries, ignoreCaseOutsideStrings);
                }
            }
            catch (Exception exception) when (exception is JsonException || exception is System.IO.IOException)
            {
                Logger.Warn(exception, "Could not read the git sync cache; starting with an empty one.");
            }

            return new SignatureCache(settings, null, ignoreCaseOutsideStrings);
        }

        /// <summary>Saves the cache. Once it grows too large only the entries used since it was loaded are kept.</summary>
        public void Save(IFileSystem fileSystem, string path = null)
        {
            path = path ?? DefaultPath;
            try
            {
                var entries = _entries.Count > MaximumEntries
                    ? _entries.Where(entry => _used.ContainsKey(entry.Key)).ToDictionary(entry => entry.Key, entry => entry.Value)
                    : _entries.ToDictionary(entry => entry.Key, entry => entry.Value);

                fileSystem.Directory.CreateDirectory(fileSystem.Path.GetDirectoryName(path));
                fileSystem.File.WriteAllText(path, JsonConvert.SerializeObject(entries));
            }
            catch (Exception exception) when (exception is System.IO.IOException || exception is UnauthorizedAccessException)
            {
                Logger.Warn(exception, "Could not save the git sync cache.");
            }
        }

        #endregion
    }
}
