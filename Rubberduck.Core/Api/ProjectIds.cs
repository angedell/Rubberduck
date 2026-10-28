using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Api
{
    /// <summary>
    /// The ids the API gives projects. The editor's own project id is a hash that can be the same for different workbooks (two files were both "㼿") and can be
    /// any character, so an agent that trusted it could write to the wrong workbook. The API's id is made from the file's path instead: letters and digits, the same
    /// for the same file, different for different files. A project that was never saved has no path; its id is made from its name and the editor's id.
    /// </summary>
    public static class ProjectIds
    {
        public static string Of(string editorId, string name, string path)
        {
            if (!string.IsNullOrWhiteSpace(path))
            {
                return "p" + Hash(path.Trim().ToLowerInvariant());
            }

            return "u" + Hash((name ?? string.Empty).ToLowerInvariant() + "|" + editorId);
        }

        /// <summary>
        /// Which projects a key addresses: the API's id first; failing that the editor's own id (what older clients stored; it may match several); failing that the name
        /// (which may match several too, the caller says so then).
        /// </summary>
        public static List<T> Match<T>(string key, IEnumerable<T> all, Func<T, (string EditorId, string Name, string Path)> describe)
        {
            var described = all.Select(p => (Item: p, Info: describe(p))).ToList();

            var byApiId = described.Where(p => string.Equals(Of(p.Info.EditorId, p.Info.Name, p.Info.Path), key, StringComparison.OrdinalIgnoreCase)).ToList();
            if (byApiId.Count > 0)
            {
                return byApiId.Select(p => p.Item).ToList();
            }

            var byEditorId = described.Where(p => string.Equals(p.Info.EditorId, key, StringComparison.OrdinalIgnoreCase)).ToList();
            if (byEditorId.Count > 0)
            {
                return byEditorId.Select(p => p.Item).ToList();
            }

            return described.Where(p => string.Equals(p.Info.Name, key, StringComparison.OrdinalIgnoreCase)).Select(p => p.Item).ToList();
        }

        /// <summary>The project's file, or null for a document that was never saved (asking for it can throw then).</summary>
        public static string PathOf(IVBProject project)
        {
            try
            {
                var name = project.FileName;
                return string.IsNullOrWhiteSpace(name) ? null : name;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return null;
            }
        }

        private static string Hash(string text)
        {
            using (var sha = SHA1.Create())
            {
                return string.Concat(sha.ComputeHash(Encoding.UTF8.GetBytes(text)).Take(4).Select(b => b.ToString("x2")));
            }
        }
    }
}
