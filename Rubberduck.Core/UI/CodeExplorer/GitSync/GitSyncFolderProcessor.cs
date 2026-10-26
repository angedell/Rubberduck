using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Text;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    public class PreparedFile
    {
        public PreparedFile(string path, string relativeDirectory, bool lineEndingsFixed, bool folderAnnotationFixed)
        {
            Path = path;
            RelativeDirectory = relativeDirectory;
            LineEndingsFixed = lineEndingsFixed;
            FolderAnnotationFixed = folderAnnotationFixed;
        }

        public string Path { get; }
        public string RelativeDirectory { get; }
        public bool LineEndingsFixed { get; }
        public bool FolderAnnotationFixed { get; }
    }

    /// <summary>
    /// Prepares the source files in a sync folder before they are imported: corrects the line endings to CRLF
    /// and makes the @Folder annotation reflect the location of the file below the sync root.
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

        public IEnumerable<string> SourceFiles(string root)
        {
            var ignored = LoadIgnoreRules(root);
            return _fileSystem.Directory
                .EnumerateFiles(root, "*", System.IO.SearchOption.AllDirectories)
                .Where(path => IsSourceFile(path) 
                               && !IsInHiddenDirectory(root, path) 
                               && !ignored.IsIgnored(RelativePath(root, path), false))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Reads every .rdsyncignore file below the root; the folders and files they exclude are not synced.
        /// </summary>
        public GitIgnoreSet LoadIgnoreRules(string root)
        {
            var files = new List<KeyValuePair<string, GitIgnore>>();
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

        private string RelativePath(string root, string path)
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

        public IEnumerable<PreparedFile> PrepareAll(string root)
        {
            return SourceFiles(root).Select(path => Prepare(root, path)).ToList();
        }

        /// <summary>
        /// Corrects the file in place if necessary.
        /// Files directly in the root are not assigned to a virtual folder, so their annotation is left alone.
        /// </summary>
        public PreparedFile Prepare(string root, string path)
        {
            var relativeDirectory = RelativeDirectory(root, path);

            var original = _fileSystem.File.ReadAllText(path, EncodingFor(path));

            var text = GitSyncText.NormalizeLineEndings(original);
            var lineEndingsFixed = text != original;

            var folderAnnotationFixed = false;
            var folderPath = GitSyncText.FolderPath(relativeDirectory);
            if (folderPath.Length > 0)
            {
                var withAnnotation = GitSyncText.SetFolderAnnotation(text, folderPath);
                folderAnnotationFixed = withAnnotation != text;
                text = withAnnotation;
            }

            if (lineEndingsFixed || folderAnnotationFixed)
            {
                _fileSystem.File.WriteAllText(path, text, EncodingFor(path));
            }

            return new PreparedFile(path, relativeDirectory, lineEndingsFixed, folderAnnotationFixed);
        }
    }
}
