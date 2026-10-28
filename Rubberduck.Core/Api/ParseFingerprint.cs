using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Api
{
    /// <summary>
    /// What a parse run reads from the editor: the content of every module and the references of every project.
    /// If it is the same as when the parser last finished, another run would change nothing and the API need not ask for one.
    /// </summary>
    public sealed class ParseFingerprint
    {
        private const int MaxNames = 5;

        private readonly Dictionary<string, int> _modules;
        private readonly string _references;

        public ParseFingerprint(IEnumerable<KeyValuePair<string, int>> moduleHashes, IEnumerable<string> references)
        {
            _modules = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var module in moduleHashes)
            {
                _modules[module.Key] = module.Value;
            }

            _referenceSet = new HashSet<string>(references, StringComparer.Ordinal);
            _references = string.Join("\n", _referenceSet.OrderBy(r => r, StringComparer.Ordinal));
        }

        private readonly HashSet<string> _referenceSet;

        public int ModuleCount => _modules.Count;

        /// <summary>True if everything in <paramref name="earlier"/> is here unchanged, with at most more modules and references besides (projects the editor did not list yet).</summary>
        public bool OnlyAddsTo(ParseFingerprint earlier)
            => earlier._referenceSet.IsSubsetOf(_referenceSet)
               && earlier._modules.All(m => _modules.TryGetValue(m.Key, out var hash) && hash == m.Value);

        /// <summary>What differs from <paramref name="earlier"/>, in words for the log; empty if nothing does.</summary>
        public IReadOnlyList<string> DifferencesFrom(ParseFingerprint earlier)
        {
            var differences = new List<string>();
            Add(differences, "added", _modules.Keys.Where(m => !earlier._modules.ContainsKey(m)));
            Add(differences, "removed", earlier._modules.Keys.Where(m => !_modules.ContainsKey(m)));
            Add(differences, "changed", _modules.Where(m => earlier._modules.TryGetValue(m.Key, out var hash) && hash != m.Value).Select(m => m.Key));
            if (_references != earlier._references)
            {
                differences.Add("references of a project changed");
            }

            return differences;
        }

        private static void Add(ICollection<string> differences, string what, IEnumerable<string> names)
        {
            var list = names.OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();
            if (list.Count == 0)
            {
                return;
            }

            var shown = string.Join(", ", list.Take(MaxNames)) + (list.Count > MaxNames ? $" and {list.Count - MaxNames} more" : string.Empty);
            differences.Add($"{what}: {shown}");
        }
    }
}
