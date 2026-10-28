using System;
using System.Linq;
using System.Security.Cryptography;

namespace Rubberduck.PackageResources
{
    public enum PackageResourceKind
    {
        /// <summary>The ribbon's XML: <c>customUI.xml</c> (Office 2007) or <c>customUI14.xml</c> (2010 and later).</summary>
        RibbonXml,

        /// <summary>The document's <c>rubberduck.json</c>: indenter and sync settings, kept in the file so they travel with the document.</summary>
        Config,

        /// <summary>A picture the ribbon XML refers to.</summary>
        Image
    }

    /// <summary>What is waiting to be written into the document's file, for a resource.</summary>
    public enum PackageResourceChange
    {
        None,
        Added,
        Modified,
        Removed
    }

    public class PackageResource
    {
        /// <summary>The file name inside the package, e.g. <c>customUI14.xml</c> or <c>logo.png</c>. Unique among the resources of a document.</summary>
        public string Name { get; set; }

        public PackageResourceKind Kind { get; set; }

        public long Size { get; set; }

        public PackageResourceChange Change { get; set; }

        /// <summary>Identifies the content as it is now (staged changes included).</summary>
        public string Hash { get; set; }

        public static string HashOf(byte[] content)
        {
            using (var sha = SHA256.Create())
            {
                return BitConverter.ToString(sha.ComputeHash(content ?? new byte[0]), 0, 16).Replace("-", string.Empty).ToLowerInvariant();
            }
        }
    }

    public class PackageResourceContent : PackageResource
    {
        public byte[] Bytes { get; set; }

        /// <summary>The content as text, for ribbon XML.</summary>
        public string Text { get; set; }
    }

    public enum PackageResourceError
    {
        NotFound,
        BadRequest,
        Conflict,
        Stale,
        HashRequired,
        Unavailable
    }

    /// <summary>What the caller got wrong, or why the document's file cannot be read right now.</summary>
    public class PackageResourceException : Exception
    {
        public PackageResourceException(PackageResourceError error, string message, Exception inner = null) : base(message, inner)
        {
            Error = error;
        }

        public PackageResourceError Error { get; }
    }

    /// <summary>The names and limits of what can live in a document's ribbon folder.</summary>
    public static class PackageResourceNames
    {
        public const string RibbonFolder = "customUI/";
        public const string ImagesFolder = "customUI/images/";
        public const string Ribbon2007 = "customUI.xml";
        public const string Ribbon2010 = "customUI14.xml";

        /// <summary>The config's name, here and in a sync folder. It is stored as <c>customUI/rubberduck.json</c>, which Excel keeps when it saves.</summary>
        public const string ConfigName = "rubberduck.json";
        public const long MaxImageBytes = 5 * 1024 * 1024;

        public static readonly string[] RibbonNames = { Ribbon2007, Ribbon2010 };

        /// <summary>The picture formats the ribbon can show, with their content types.</summary>
        public static readonly string[] ImageExtensions = { ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".ico" };

        //Only the documents Office keeps as a zip of parts. Access and the binary formats (.xls, .doc) have no ribbon folder.
        public static readonly string[] DocumentExtensions =
        {
            ".xlsm", ".xlsx", ".xlam", ".xltm", ".xltx",
            ".docm", ".docx", ".dotm", ".dotx",
            ".pptm", ".pptx", ".potm", ".ppam", ".ppsm"
        };

        public static bool IsConfigName(string name) => string.Equals(name, ConfigName, StringComparison.OrdinalIgnoreCase);

        public static PackageResourceKind KindOf(string name)
            => IsRibbonName(name) ? PackageResourceKind.RibbonXml : IsConfigName(name) ? PackageResourceKind.Config : PackageResourceKind.Image;

        public static bool IsRibbonName(string name) => RibbonNames.Contains(name, StringComparer.OrdinalIgnoreCase);

        public static bool IsImageName(string name)
        {
            var extension = System.IO.Path.GetExtension(name ?? string.Empty);
            return IsValidFileName(name) && ImageExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase);
        }

        //Starts with a letter or underscore, then letters, digits, underscore, dash, dot: safe as a zip entry, and valid as the relationship id the ribbon XML's 'image' attribute names.
        public static bool IsValidFileName(string name)
            => !string.IsNullOrEmpty(name) && name.Length <= 100 && System.Text.RegularExpressions.Regex.IsMatch(name, @"^[A-Za-z_][A-Za-z0-9_.\-]*$");

        public static string ContentTypeOf(string extension)
        {
            switch ((extension ?? string.Empty).ToLowerInvariant())
            {
                case ".png": return "image/png";
                case ".jpg":
                case ".jpeg": return "image/jpeg";
                case ".gif": return "image/gif";
                case ".bmp": return "image/bmp";
                case ".ico": return "image/x-icon";
                case ".json": return "application/json";
                default: return "application/xml";
            }
        }
    }
}
