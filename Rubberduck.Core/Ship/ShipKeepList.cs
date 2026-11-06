using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace Rubberduck.Ship
{
    /// <summary>The one check of the tree shaking keep list, used by the settings file and by the API.</summary>
    public static class ShipKeepList
    {
        public const int MaxEntries = 500;
        public const int MaxLength = 200;

        public const string KeepNotAnArray = "'treeShaking.keep' must be an array of names (strings), \"Member\" or \"Module.Member\".";
        public const string ModeInvalid = "'treeShaking.mode' must be \"conservative\" or \"entryPointsOnly\".";

        /// <summary>
        /// The one reading of the keep list, for the settings file and the API: entries are trimmed; an item that is not a string is an error for that item only;
        /// entries beyond the first <see cref="MaxEntries"/> are an error and are not looked at. Valid is null if <paramref name="keep"/> is not an array.
        /// </summary>
        public static (List<string> Valid, List<string> Errors) Normalize(JToken keep)
        {
            var errors = new List<string>();
            if (!(keep is JArray array))
            {
                errors.Add(KeepNotAnArray);
                return (null, errors);
            }

            var valid = new List<string>();
            for (var i = 0; i < array.Count; i++)
            {
                if (i >= MaxEntries)
                {
                    errors.Add($"'treeShaking.keep' holds at most {MaxEntries} names; the others are ignored.");
                    break;
                }

                if (array[i].Type != JTokenType.String)
                {
                    errors.Add($"Entry {i + 1} of 'treeShaking.keep' must be a string.");
                    continue;
                }

                valid.AddRange(Validate(new[] { array[i].Value<string>() }, errors));
            }

            return (valid, errors);
        }

        /// <summary>The mode, trimmed and without regard to case; null with the error if it is not one of the two.</summary>
        public static (ShipTreeShakingMode? Mode, string Error) NormalizeMode(JToken mode)
        {
            var text = mode != null && mode.Type == JTokenType.String ? mode.Value<string>().Trim() : null;
            if (string.Equals(text, "conservative", System.StringComparison.OrdinalIgnoreCase)) return (ShipTreeShakingMode.Conservative, null);
            if (string.Equals(text, "entryPointsOnly", System.StringComparison.OrdinalIgnoreCase)) return (ShipTreeShakingMode.EntryPointsOnly, null);
            return (null, ModeInvalid);
        }

        /// <summary>
        /// Returns the entries that are fine (trimmed), in order. Every entry that is not (empty, too long, a control character, or beyond the
        /// first <see cref="MaxEntries"/>) is left out and described in <paramref name="errors"/>.
        /// </summary>
        public static List<string> Validate(IEnumerable<string> names, ICollection<string> errors)
        {
            var valid = new List<string>();
            var count = 0;
            foreach (var raw in names)
            {
                count++;
                if (count > MaxEntries)
                {
                    if (count == MaxEntries + 1)
                    {
                        errors.Add($"'treeShaking.keep' holds at most {MaxEntries} names; the others are ignored.");
                    }

                    continue;
                }

                var name = (raw ?? string.Empty).Trim();
                if (name.Length == 0)
                {
                    errors.Add("A name in 'treeShaking.keep' is empty.");
                }
                else if (name.Length > MaxLength)
                {
                    errors.Add($"A name in 'treeShaking.keep' is longer than {MaxLength} characters.");
                }
                else if (name.Any(char.IsControl))
                {
                    errors.Add("A name in 'treeShaking.keep' has a control character.");
                }
                else
                {
                    valid.Add(name);
                }
            }

            return valid;
        }
    }
}
