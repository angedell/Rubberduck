using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace Rubberduck.PackageResources
{
    /// <summary>What is wrong with a piece of text (ribbon XML, rubberduck.json), and where.</summary>
    public class TextProblem
    {
        public TextProblem(string message, int line = 0, int column = 0, bool blocking = true)
        {
            Message = message;
            Line = line;
            Column = column;
            Blocking = blocking;
        }

        public string Message { get; }

        /// <summary>1-based; 0 when the problem is not tied to a place.</summary>
        public int Line { get; }

        public int Column { get; }

        /// <summary>False for a warning: the text can still be saved.</summary>
        public bool Blocking { get; }
    }

    /// <summary>Checks and lays out the XML of <c>customUI.xml</c> / <c>customUI14.xml</c>.</summary>
    public static class RibbonXml
    {
        public const string Namespace2007 = "http://schemas.microsoft.com/office/2006/01/customui";
        public const string Namespace2010 = "http://schemas.microsoft.com/office/2009/07/customui";

        public static string NamespaceOf(string name)
            => name.Equals(PackageResourceNames.Ribbon2007, StringComparison.OrdinalIgnoreCase) ? Namespace2007 : Namespace2010;

        /// <summary>Null if Office will accept the XML as that part: well-formed, root <c>customUI</c>, in the namespace its file name belongs to.</summary>
        public static TextProblem Check(string name, string xml)
        {
            if (string.IsNullOrWhiteSpace(xml))
            {
                return new TextProblem("The XML is empty.");
            }

            XDocument document;
            try
            {
                document = XDocument.Parse(xml, LoadOptions.SetLineInfo);
            }
            catch (XmlException exception)
            {
                return new TextProblem(exception.Message, exception.LineNumber, exception.LinePosition);
            }

            var root = document.Root;
            if (root.Name.LocalName != "customUI")
            {
                return new TextProblem($"The root element must be <customUI>, not <{root.Name.LocalName}>.", 1, 1);
            }

            var expected = NamespaceOf(name);
            if (root.Name.NamespaceName != expected)
            {
                return new TextProblem($"{name} needs <customUI xmlns=\"{expected}\">; Office ignores the ribbon otherwise.", 1, 1);
            }

            //Office refuses a ribbon in which two controls share an id, and says little about which.
            var seen = new System.Collections.Generic.Dictionary<string, XElement>();
            foreach (var element in root.DescendantsAndSelf())
            {
                var id = (string)element.Attribute("id");
                if (id == null)
                {
                    continue;
                }

                if (seen.TryGetValue(id, out _))
                {
                    var info = (IXmlLineInfo)element;
                    return new TextProblem($"The id \"{id}\" is used twice; every control needs its own.", info.HasLineInfo() ? info.LineNumber : 0, info.HasLineInfo() ? info.LinePosition : 0);
                }

                seen.Add(id, element);
            }

            return null;
        }

        /// <summary>The XML indented, with the declaration and comments kept. Throws <see cref="XmlException"/> if it is not well-formed.</summary>
        public static string Format(string xml) => Write(xml, true);

        /// <summary>The XML on a single line: the layout whitespace between elements is dropped, the declaration and comments kept. Throws <see cref="XmlException"/> if it is not well-formed.</summary>
        public static string Linearize(string xml) => Write(xml, false);

        private static string Write(string xml, bool indent)
        {
            var document = XDocument.Parse(xml, LoadOptions.None);
            var settings = new XmlWriterSettings { Indent = indent, IndentChars = "  ", OmitXmlDeclaration = true, NewLineChars = "\r\n", NewLineHandling = NewLineHandling.Replace };
            var builder = new StringBuilder();
            using (var writer = XmlWriter.Create(builder, settings))
            {
                document.Save(writer);
            }

            var declaration = Regex.Match(xml, @"^\s*<\?xml[^>]*\?>");
            return declaration.Success
                ? declaration.Value.Trim() + (indent ? "\r\n" : string.Empty) + builder
                : builder.ToString();
        }

        /// <summary>The text as it is stored: UTF-8, and a declaration that says so.</summary>
        public static byte[] ToBytes(string xml)
        {
            var text = Regex.Replace(xml, @"^(\s*<\?xml[^>]*?)encoding\s*=\s*(""[^""]*""|'[^']*')", "$1encoding=\"UTF-8\"", RegexOptions.IgnoreCase);
            return new UTF8Encoding(false).GetBytes(text);
        }

        /// <summary>A minimal ribbon to start from: one tab with one button.</summary>
        public static string Template(string name) => string.Join("\r\n",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
            "<customUI xmlns=\"" + NamespaceOf(name) + "\">",
            "  <ribbon>",
            "    <tabs>",
            "      <tab id=\"customTab\" label=\"My Tab\">",
            "        <group id=\"customGroup\" label=\"My Group\">",
            "          <button id=\"customButton\" label=\"Run\" size=\"large\" imageMso=\"HappyFace\" onAction=\"OnAction\" />",
            "        </group>",
            "      </tab>",
            "    </tabs>",
            "  </ribbon>",
            "</customUI>",
            string.Empty);
    }
}
