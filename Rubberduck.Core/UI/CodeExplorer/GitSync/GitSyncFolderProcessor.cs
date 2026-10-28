using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Text;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>
    /// Finds the source files in a sync folder and the rules that say which of them are left alone. Reads only.
    /// </summary>
    public class GitSyncFolderProcessor
    {
        private static readonly string[] SourceExtensions =
        {
            ComponentTypeExtensions.StandardExtension,
            ComponentTypeExtensions.ClassExtension,
            ComponentTypeExtensions.FormExtension,
            ComponentTypeExtensions.DocClassExtension
        };

        private readonly IFileSystem _fileSystem;

        public GitSyncFolderProcessor(IFileSystem fileSystem)
        {
            _fileSystem = fileSystem;
        }

        /// <summary>
        /// Name of the files listing what not to sync. They use the .gitignore syntax but are separate from git's own.
        /// </summary>
        public const string IgnoreFileName = ".rdsyncignore";

        //The VBE reads and writes source files in the ANSI code page of the current user.
        public static Encoding SourceEncoding => Encoding.Default;

        /// <summary>
        /// Document modules are exported as UTF-8, because they are the code pane's text and not a file the VBE itself reads back.
        /// </summary>
        public static Encoding EncodingFor(string path)
        {
            return System.IO.Path.GetExtension(path).Equals(ComponentTypeExtensions.DocClassExtension, StringComparison.OrdinalIgnoreCase)
                ? new UTF8Encoding(false)
                : SourceEncoding;
        }

        public bool IsSourceFile(string path)
        {
            var extension = _fileSystem.Path.GetExtension(path);
            return SourceExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>Every source file below the root, except those in hidden folders such as .git. Ignore rules are not applied.</summary>
        public IEnumerable<string> AllSourceFiles(string root)
        {
            return _fileSystem.Directory
                .EnumerateFiles(root, "*", System.IO.SearchOption.AllDirectories)
                .Where(path => IsSourceFile(path) && !IsInHiddenDirectory(root, path))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>Every markdown file below the root, except those in hidden folders such as .git. Ignore rules are not applied.</summary>
        public IEnumerable<string> AllNoteFiles(string root)
        {
            return _fileSystem.Directory
                .EnumerateFiles(root, "*.md", System.IO.SearchOption.AllDirectories)
                .Where(path => _fileSystem.Path.GetExtension(path).Equals(".md", StringComparison.OrdinalIgnoreCase) && !IsInHiddenDirectory(root, path))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase);
        }

        public IEnumerable<string> SourceFiles(string root)
        {
            var ignored = LoadIgnoreRules(root);
            return AllSourceFiles(root).Where(path => !ignored.IsIgnored(RelativePath(root, path), false));
        }

        /// <summary>
        /// Reads every .rdsyncignore file below the root, and adds the patterns of the config's <c>sync.ignore</c> as if they were in a file in the root.
        /// The folders and files they exclude are not synced.
        /// </summary>
        public GitIgnoreSet LoadIgnoreRules(string root, IEnumerable<string> configured = null)
        {
            var files = new List<KeyValuePair<string, GitIgnore>>();
            if (configured != null)
            {
                files.Add(new KeyValuePair<string, GitIgnore>(string.Empty, new GitIgnore(configured.ToList())));
            }

            foreach (var path in _fileSystem.Directory.EnumerateFiles(root, IgnoreFileName, System.IO.SearchOption.AllDirectories))
            {
                if (IsInHiddenDirectory(root, path))
                {
                    continue;
                }

                var folder = RelativeDirectory(root, path).Replace('\\', '/');
                files.Add(new KeyValuePair<string, GitIgnore>(folder, new GitIgnore(_fileSystem.File.ReadAllLines(path))));
            }

            return new GitIgnoreSet(files);
        }

        public bool IsIgnored(string root, string path)
        {
            return LoadIgnoreRules(root).IsIgnored(RelativePath(root, path), false);
        }

        public string RelativePath(string root, string path)
        {
            var directory = RelativeDirectory(root, path);
            var name = _fileSystem.Path.GetFileName(path);
            return directory.Length == 0 ? name : directory + "\\" + name;
        }

        //Skips .git and the like.
        private bool IsInHiddenDirectory(string root, string path)
        {
            return RelativeDirectory(root, path)
                .Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries)
                .Any(segment => segment.StartsWith("."));
        }

        public string RelativeDirectory(string root, string path)
        {
            var rootPath = _fileSystem.Path.GetFullPath(root).TrimEnd('\\', '/') + "\\";
            var directory = _fileSystem.Path.GetDirectoryName(_fileSystem.Path.GetFullPath(path)) + "\\";

            return directory.StartsWith(rootPath, StringComparison.OrdinalIgnoreCase)
                ? directory.Substring(rootPath.Length).TrimEnd('\\')
                : string.Empty;
        }
    }
}
