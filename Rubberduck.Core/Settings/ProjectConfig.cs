using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Reflection;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using Rubberduck.SmartIndenter;

namespace Rubberduck.Settings
{
    /// <summary>
    /// What was read from a <c>rubberduck.json</c>: the indenter settings it asks for, plus anything wrong with it.
    /// </summary>
    public class ProjectConfig
    {
        public const string FileName = "rubberduck.json";

        /// <summary>The path of the file this was read from, or null if there is none.</summary>
        public string Path { get; set; }

        /// <summary>The indenter settings, with the file's values applied over the base settings. Null if there is no file.</summary>
        public IIndenterSettings Indenter { get; set; }

        /// <summary>What the file says about syncing the project with its folder. Empty if there is no file.</summary>
        public SyncConfig Sync { get; } = new SyncConfig();

        /// <summary>Unknown keys, bad values and malformed JSON. The file is ignored for whatever could not be read.</summary>
        public List<string> Errors { get; } = new List<string>();

        public bool Found => Path != null;
    }

    public enum ViolationMode
    {
        /// <summary>Modules and files that break an enforced rule cannot be synced.</summary>
        Block,
        /// <summary>They are flagged but can still be synced.</summary>
        Warn,
        /// <summary>They are repaired as they are synced: missing annotations and attributes are added, attributes are set to what their annotation says.</summary>
        Fix
    }

    public class SyncConfig
    {
        /// <summary>The sync rules that can be enforced.</summary>
        public static readonly string[] Rules =
        {
            "AttributeValueOutOfSync",
            "MissingAttribute",
            "MissingModuleAnnotation",
            "ModuleWithoutFolder"
        };

        /// <summary>Patterns, in the syntax of .gitignore and relative to the sync folder, of the files that are never synced.</summary>
        public List<string> Ignore { get; } = new List<string>();

        /// <summary>The names of the rules that modules and files have to follow to be synced.</summary>
        public List<string> Enforce { get; } = new List<string>();

        public ViolationMode OnViolation { get; set; } = ViolationMode.Fix;

        /// <summary>
        /// VBA does not care about case, so a difference in case outside of string literals is not a difference. Inside a string it is.
        /// </summary>
        public bool IgnoreCaseOutsideStrings { get; set; } = true;
    }

    /// <summary>
    /// Reads the repository-level config file, in the way Biome reads <c>biome.json</c>: the nearest file in the sync folder
    /// or any of its parents wins, keys that are left out keep the user's own Rubberduck settings, and unknown keys are reported rather than ignored.
    /// <code>
    /// {
    ///   "indenter": {
    ///     "indentSpaces": 2,
    ///     "indentCase": true,
    ///     "endOfLineCommentStyle": "SameGap"
    ///   },
    ///   "sync": {
    ///     "ignore": ["modScratch.bas", "Sandbox/"],
    ///     "enforce": ["ModuleWithoutFolder", "MissingAttribute"],
    ///     "onViolation": "fix",
    ///     "ignoreCaseOutsideStrings": true
    ///   }
    /// }
    /// </code>
    /// </summary>
    public class ProjectConfigReader
    {
        private const string IndenterKey = "indenter";
        private const string SyncKey = "sync";

        // The four settings that overlap each other, in the order that gives the intended result when copied one by one.
        private static readonly string[] CopiedFirst =
        {
            nameof(IIndenterSettings.ForceDebugStatementsInColumn1),
            nameof(IIndenterSettings.ForceDebugPrintInColumn1),
            nameof(IIndenterSettings.ForceDebugAssertInColumn1),
            nameof(IIndenterSettings.ForceStopInColumn1)
        };

        private static readonly PropertyInfo[] IndenterProperties = typeof(IndenterSettings)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead && p.CanWrite && p.DeclaringType == typeof(IndenterSettings))
            .OrderBy(p => Array.IndexOf(CopiedFirst, p.Name) is var i && i >= 0 ? i : CopiedFirst.Length)
            .ToArray();

        private readonly IFileSystem _fileSystem;

        public ProjectConfigReader(IFileSystem fileSystem)
        {
            _fileSystem = fileSystem;
        }

        /// <param name="folder">Where to start looking; parents are searched too.</param>
        /// <param name="baseSettings">The settings to use for everything the file does not mention.</param>
        public ProjectConfig Read(string folder, IIndenterSettings baseSettings)
        {
            var config = new ProjectConfig();
            var path = Find(folder);
            if (path == null)
            {
                return config;
            }

            config.Path = path;
            var settings = Copy(baseSettings);

            JObject root;
            try
            {
                root = JObject.Parse(_fileSystem.File.ReadAllText(path));
            }
            catch (Exception exception) when (exception is JsonException || exception is System.IO.IOException)
            {
                config.Errors.Add($"{ProjectConfig.FileName} could not be read: {exception.Message}");
                config.Indenter = settings;
                return config;
            }

            foreach (var section in root.Properties())
            {
                if (section.Name.Equals("$schema", StringComparison.Ordinal))
                {
                    continue;
                }

                var isIndenter = section.Name.Equals(IndenterKey, StringComparison.OrdinalIgnoreCase);
                var isSync = section.Name.Equals(SyncKey, StringComparison.OrdinalIgnoreCase);
                if (!isIndenter && !isSync)
                {
                    config.Errors.Add($"Unknown section '{section.Name}'.");
                    continue;
                }

                if (!(section.Value is JObject body))
                {
                    config.Errors.Add($"'{section.Name}' must be an object.");
                    continue;
                }

                if (isIndenter)
                {
                    Apply(body, settings, config.Errors);
                }
                else
                {
                    ReadSync(body, config.Sync, config.Errors);
                }
            }

            config.Indenter = settings;
            return config;
        }

        private static void ReadSync(JObject section, SyncConfig sync, ICollection<string> errors)
        {
            foreach (var property in section.Properties())
            {
                switch (property.Name.ToLowerInvariant())
                {
                    case "ignore":
                        sync.Ignore.AddRange(Strings(property, errors));
                        break;
                    case "enforce":
                        foreach (var rule in Strings(property, errors))
                        {
                            var known = SyncConfig.Rules.FirstOrDefault(r => r.Equals(rule, StringComparison.OrdinalIgnoreCase));
                            if (known == null)
                            {
                                errors.Add($"Unknown sync rule '{rule}'. Known rules: {string.Join(", ", SyncConfig.Rules)}.");
                            }
                            else if (!sync.Enforce.Contains(known))
                            {
                                sync.Enforce.Add(known);
                            }
                        }
                        break;
                    case "onviolation":
                        if (Enum.TryParse(property.Value.Type == JTokenType.String ? property.Value.Value<string>() : null, true, out ViolationMode mode)
                            && Enum.IsDefined(typeof(ViolationMode), mode))
                        {
                            sync.OnViolation = mode;
                        }
                        else
                        {
                            errors.Add($"Invalid value for sync setting 'onViolation': {property.Value}. Use \"fix\", \"block\" or \"warn\".");
                        }
                        break;
                    case "ignorecaseoutsidestrings":
                        if (property.Value.Type == JTokenType.Boolean)
                        {
                            sync.IgnoreCaseOutsideStrings = property.Value.Value<bool>();
                        }
                        else
                        {
                            errors.Add($"Invalid value for sync setting 'ignoreCaseOutsideStrings': {property.Value}. Use true or false.");
                        }
                        break;
                    case "$comment":
                        break;
                    default:
                        errors.Add($"Unknown sync setting '{property.Name}'.");
                        break;
                }
            }
        }

        private static IEnumerable<string> Strings(JProperty property, ICollection<string> errors)
        {
            if (!(property.Value is JArray array) || array.Any(item => item.Type != JTokenType.String))
            {
                errors.Add($"Sync setting '{property.Name}' must be a list of strings.");
                return Enumerable.Empty<string>();
            }

            return array.Select(item => item.Value<string>()).Where(item => !string.IsNullOrWhiteSpace(item)).ToList();
        }

        public string Find(string folder)
        {
            if (string.IsNullOrWhiteSpace(folder))
            {
                return null;
            }

            var directory = _fileSystem.DirectoryInfo.FromDirectoryName(folder);
            while (directory != null)
            {
                var candidate = _fileSystem.Path.Combine(directory.FullName, ProjectConfig.FileName);
                if (_fileSystem.File.Exists(candidate))
                {
                    return candidate;
                }

                directory = directory.Parent;
            }

            return null;
        }

        private static void Apply(JObject section, IndenterSettings settings, ICollection<string> errors)
        {
            var serializer = JsonSerializer.Create(new JsonSerializerSettings { Converters = { new StringEnumConverter() } });

            // Apply in the order of IndenterProperties, so that e.g. forceDebugStatementsInColumn1 never overrides forceStopInColumn1 from the same file.
            var requested = section.Properties().ToList();
            foreach (var property in requested)
            {
                if (!IndenterProperties.Any(p => p.Name.Equals(property.Name, StringComparison.OrdinalIgnoreCase)))
                {
                    errors.Add($"Unknown indenter setting '{property.Name}'.");
                }
            }

            foreach (var target in IndenterProperties)
            {
                var property = requested.FirstOrDefault(p => p.Name.Equals(target.Name, StringComparison.OrdinalIgnoreCase));
                if (property == null)
                {
                    continue;
                }

                try
                {
                    var value = property.Value.ToObject(target.PropertyType, serializer);
                    if (value == null || (target.PropertyType.IsEnum && !Enum.IsDefined(target.PropertyType, value)))
                    {
                        throw new JsonException("not a valid value");
                    }

                    target.SetValue(settings, value);
                }
                catch (Exception exception) when (exception is JsonException || exception is ArgumentException || exception is FormatException)
                {
                    errors.Add($"Invalid value for indenter setting '{target.Name}': {property.Value}.");
                }
            }
        }

        /// <summary>A copy of the settings that can be changed without touching the user's own.</summary>
        public static IndenterSettings Copy(IIndenterSettings source)
        {
            var copy = new IndenterSettings(true);
            if (source == null)
            {
                return copy;
            }

            foreach (var property in IndenterProperties)
            {
                var value = typeof(IIndenterSettings).GetProperty(property.Name)?.GetValue(source);
                if (value != null)
                {
                    property.SetValue(copy, value);
                }
            }

            return copy;
        }
    }
}
