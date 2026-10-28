using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>Prints a note as the preview shows it (rendered markdown, in the chosen font), not as the code of its module.</summary>
    public static class NotePrinter
    {
        private const double MarginInDevicePixels = 48;

        /// <summary>Asks which printer and prints. False if the user cancelled.</summary>
        public static bool Print(string markdown, string documentName, string fontName)
        {
            var dialog = new PrintDialog();
            if (dialog.ShowDialog() != true)
            {
                return false;
            }

            var document = MarkdownRenderer.ToDocument(markdown, fontName);
            document.PageWidth = dialog.PrintableAreaWidth;
            document.PageHeight = dialog.PrintableAreaHeight;
            document.PagePadding = new Thickness(MarginInDevicePixels);
            //One column the width of the page, as on paper.
            document.ColumnWidth = dialog.PrintableAreaWidth;
            document.IsOptimalParagraphEnabled = false;

            dialog.PrintDocument(((IDocumentPaginatorSource)document).DocumentPaginator, documentName);
            return true;
        }
    }
}
