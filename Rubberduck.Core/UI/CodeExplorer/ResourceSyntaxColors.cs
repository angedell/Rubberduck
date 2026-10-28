using System.Xml;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>Syntax colours for the resource editor. XML uses AvalonEdit's own definition; JSON has none there, so it is described here.</summary>
    public static class ResourceSyntaxColors
    {
        //No spans: JSON strings never run over a line, and one ordered list of rules lets the key rule win over the plain string rule.
        private const string JsonDefinition = @"<?xml version=""1.0""?>
<SyntaxDefinition name=""Json"" xmlns=""http://icsharpcode.net/sharpdevelop/syntaxdefinition/2008"">
  <Color name=""Key"" foreground=""#0451A5"" />
  <Color name=""String"" foreground=""#A31515"" />
  <Color name=""Number"" foreground=""#098658"" />
  <Color name=""Literal"" foreground=""#0000FF"" />
  <Color name=""Comment"" foreground=""#008000"" />
  <RuleSet>
    <Rule color=""Comment"">//.*</Rule>
    <Rule color=""Key"">""(?:[^""\\\r\n]|\\.)*""(?=\s*:)</Rule>
    <Rule color=""String"">""(?:[^""\\\r\n]|\\.)*""</Rule>
    <Rule color=""Number"">-?\b\d+(?:\.\d+)?(?:[eE][+-]?\d+)?\b</Rule>
    <Rule color=""Literal"">\b(?:true|false|null)\b</Rule>
  </RuleSet>
</SyntaxDefinition>";

        private static IHighlightingDefinition _json;

        public static IHighlightingDefinition Xml => HighlightingManager.Instance.GetDefinition("XML");

        public static IHighlightingDefinition Json
        {
            get
            {
                if (_json == null)
                {
                    using (var reader = XmlReader.Create(new System.IO.StringReader(JsonDefinition)))
                    {
                        _json = HighlightingLoader.Load(reader, HighlightingManager.Instance);
                    }
                }

                return _json;
            }
        }
    }
}
