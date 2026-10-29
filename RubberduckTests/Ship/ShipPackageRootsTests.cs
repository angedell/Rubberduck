using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Ship;

namespace RubberduckTests.Ship
{
    /// <summary>Document roots (ADR 0004 item 3): names found in the parts of the document's package. A pure function of the parts' text.</summary>
    [TestFixture]
    [Category("Ship")]
    public class ShipPackageRootsTests
    {
        private static readonly string[] Names = { "OnRibbonClick", "GetLabel", "ClickMe", "MyUdf", "Report", "Unused" };

        private static List<PackageRootHit> Find(string part, string text, params string[] names)
            => PackageRoots.Find(new Dictionary<string, string> { { part, text } }, names.Length == 0 ? Names : names);

        private static MemoryStream Zip(params (string Path, string Text)[] parts)
        {
            var stream = new MemoryStream();
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, true))
            {
                foreach (var (path, text) in parts)
                {
                    using (var writer = new StreamWriter(zip.CreateEntry(path).Open()))
                    {
                        writer.Write(text);
                    }
                }
            }

            stream.Position = 0;
            return stream;
        }

        [Test]
        public void RibbonCallbackAttributesAreRoots()
        {
            var xml = "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><ribbon><tabs><tab id=\"t1\"><group id=\"g1\" label=\"A label\">"
                      + "<button id=\"b1\" onAction=\"OnRibbonClick\" getLabel=\"GetLabel\"/></group></tab></tabs></ribbon></customUI>";

            var hits = Find("customUI/customUI14.xml", xml);

            CollectionAssert.AreEquivalent(new[] { "OnRibbonClick", "GetLabel" }, hits.Select(h => h.Name).ToList());
            StringAssert.Contains("onAction", hits.Single(h => h.Name == "OnRibbonClick").Source);
        }

        [Test]
        public void RibbonTextThatIsNotAnAttributeValueIsNotARoot()
        {
            var xml = "<customUI><ribbon><tabs><tab id=\"t1\"><group id=\"g1\"><button id=\"b1\" label=\"Not an identifier\"/>Unused</group></tab></tabs></ribbon></customUI>";

            Assert.IsEmpty(Find("customUI/customUI.xml", xml));
        }

        [TestCase("Module1.OnRibbonClick")]
        [TestCase("'Book.xlam'!OnRibbonClick")]
        [TestCase("Module1.OnRibbonClick 1")]
        public void ARibbonAttributeValueThatIsNotABareIdentifierStillRootsEveryWordInIt(string value)
        {
            var xml = "<customUI><ribbon><tabs><tab id=\"t1\"><group id=\"g1\"><button id=\"b1\" onAction=\"" + value.Replace("'", "&apos;") + "\"/></group></tab></tabs></ribbon></customUI>";

            var hits = Find("customUI/customUI14.xml", xml);

            CollectionAssert.Contains(hits.Select(h => h.Name).ToList(), "OnRibbonClick");
        }

        [TestCase("xl/tables/table1.xml", "<table><tableColumns><tableColumn id=\"1\"><calculatedColumnFormula>MyUdf([@Qty])</calculatedColumnFormula></tableColumn></tableColumns></table>")]
        [TestCase("xl/pivotCache/pivotCacheDefinition1.xml", "<pivotCacheDefinition><cacheFields><cacheField name=\"x\" formula=\"MyUdf(Qty)\"/></cacheFields></pivotCacheDefinition>")]
        [TestCase("xl/pivotCache2/pivotCacheDefinition2.xml", "<pivotCacheDefinition><cacheFields><cacheField name=\"x\" formula=\"MyUdf(Qty)\"/></cacheFields></pivotCacheDefinition>")]
        [TestCase("xl/pivotTables/pivotTable1.xml", "<pivotTableDefinition><calculatedItems><calculatedItem formula=\"MyUdf(Qty)\"/></calculatedItems></pivotTableDefinition>")]
        [TestCase("xl/queryTables/queryTable1.xml", "<queryTable><queryTableRefresh><queryTableFields><queryTableField id=\"1\" name=\"MyUdf\"/></queryTableFields></queryTableRefresh></queryTable>")]
        public void FormulasOfTablesPivotsAndQueryTablesAreSearchedToo(string part, string xml)
        {
            var hits = PackageRoots.Find(Zip((part, xml)), Names);

            CollectionAssert.AreEqual(new[] { "MyUdf" }, hits.Select(h => h.Name).ToList());
            Assert.AreEqual(part, hits.Single().Source);
        }

        [Test]
        public void TheListOfPartsThatAreReadIsADecision()
        {
            var read = new[]
            {
                "customUI/customUI.xml", "customUI/customUI14.xml",
                "xl/workbook.xml", "xl/connections.xml",
                "xl/worksheets/sheet1.xml", "xl/chartsheets/sheet2.xml", "xl/macrosheets/sheet3.xml", "xl/dialogsheets/sheet4.xml",
                "xl/drawings/drawing1.xml", "xl/drawings/vmlDrawing1.vml", "xl/charts/chart1.xml", "xl/ctrlProps/ctrlProp1.xml",
                "xl/tables/table1.xml", "xl/pivotTables/pivotTable1.xml", "xl/pivotCache/pivotCacheDefinition1.xml", "xl/queryTables/queryTable1.xml",
                "xl/slicerCaches/slicerCache1.xml", "xl/slicers/slicer1.xml", "xl/externalLinks/externalLink1.xml"
            };
            var notRead = new[]
            {
                "xl/sharedStrings.xml", "xl/styles.xml", "xl/pivotCache/pivotCacheRecords1.xml", "xl/vbaProject.bin", "docProps/core.xml", "[Content_Types].xml", "xl/theme/theme1.xml"
            };

            foreach (var part in read) Assert.IsTrue(PackageRoots.IsRelevantPart(part), part + " should be read");
            foreach (var part in notRead) Assert.IsFalse(PackageRoots.IsRelevantPart(part), part + " should not be read");
        }

        [TestCase("xl/connections.xml")]
        [TestCase("xl/slicerCaches/slicerCache1.xml")]
        [TestCase("xl/slicers/slicer1.xml")]
        [TestCase("xl/externalLinks/externalLink1.xml")]
        public void MacroNamesInConnectionsSlicersAndExternalLinksAreRoots(string part)
        {
            var hits = PackageRoots.Find(Zip((part, "<root><thing macro=\"MyUdf\"/></root>")), Names);

            CollectionAssert.AreEqual(new[] { "MyUdf" }, hits.Select(h => h.Name).ToList());
        }

        [Test]
        public void ThePivotCacheRecordsAreDataAndNotSearched()
        {
            var hits = PackageRoots.Find(Zip(("xl/pivotCache/pivotCacheRecords1.xml", "<pivotCacheRecords><r><s v=\"MyUdf\"/></r></pivotCacheRecords>")), Names);

            Assert.IsEmpty(hits);
        }

        [Test]
        public void ShapeMacrosOnWorksheetsAndDrawingsAreRoots()
        {
            var drawing = "<xdr:wsDr xmlns:xdr=\"d\"><xdr:sp macro=\"[0]!ClickMe\" textlink=\"\"/></xdr:wsDr>";
            var sheet = "<worksheet><sheetData><c r=\"A1\"><f>MyUdf(B1)+1</f><v>3</v></c><c r=\"A2\" t=\"inlineStr\"><is><t>Report</t></is></c></sheetData></worksheet>";

            var fromDrawing = Find("xl/drawings/drawing1.xml", drawing);
            var fromSheet = Find("xl/worksheets/sheet1.xml", sheet);

            CollectionAssert.AreEqual(new[] { "ClickMe" }, fromDrawing.Select(h => h.Name).ToList());
            CollectionAssert.AreEqual(new[] { "MyUdf" }, fromSheet.Select(h => h.Name).ToList(), "A cell's value is data; its formula is code.");
        }

        [Test]
        public void DefinedNamesAndChartMacrosAreRoots()
        {
            var workbook = "<workbook><definedNames><definedName name=\"Calc\">Module1.Report()</definedName></definedNames></workbook>";
            var chart = "<c:chartSpace xmlns:c=\"c\"><c:userShapes macro=\"ClickMe\"/></c:chartSpace>";

            CollectionAssert.AreEqual(new[] { "Report" }, Find("xl/workbook.xml", workbook).Select(h => h.Name).ToList());
            CollectionAssert.AreEqual(new[] { "ClickMe" }, Find("xl/charts/chart1.xml", chart).Select(h => h.Name).ToList());
        }

        [Test]
        public void ANameIsMatchedAsAWholeWordNotAsPartOfAnother()
        {
            var sheet = "<worksheet><sheetData><c r=\"A1\"><f>MyUdf2(B1)+XMyUdf(1)+MyUdf_x(2)</f></c></sheetData></worksheet>";

            Assert.IsEmpty(Find("xl/worksheets/sheet1.xml", sheet));
        }

        [Test]
        public void NamesAreMatchedWithoutRegardToCase()
        {
            var sheet = "<worksheet><sheetData><c r=\"A1\"><f>myudf(B1)</f></c></sheetData></worksheet>";

            CollectionAssert.AreEqual(new[] { "MyUdf" }, Find("xl/worksheets/sheet1.xml", sheet).Select(h => h.Name).ToList());
        }

        [Test]
        public void PartsThatCannotHoldMacroNamesAreNotRead()
        {
            Assert.IsEmpty(Find("xl/sharedStrings.xml", "<sst><si><t>MyUdf</t></si></sst>"));
            Assert.IsEmpty(Find("docProps/core.xml", "<cp>MyUdf</cp>"));
        }

        [Test]
        public void MalformedXmlIsStillSearchedAsText()
        {
            var hits = Find("xl/drawings/vmlDrawing1.vml", "<xml><x:ClientData><x:FmlaMacro>[0]!ClickMe</x:FmlaMacro><br></x:ClientData>");

            CollectionAssert.AreEqual(new[] { "ClickMe" }, hits.Select(h => h.Name).ToList());
        }

        [Test]
        public void ReadsTheParthsOutOfAZipPackage()
        {
            using (var package = Zip(
                       ("[Content_Types].xml", "<Types/>"),
                       ("customUI/customUI14.xml", "<customUI><button onAction=\"OnRibbonClick\"/></customUI>"),
                       ("xl/worksheets/sheet1.xml", "<worksheet><sheetData><c><f>MyUdf(1)</f></c></sheetData></worksheet>"),
                       ("xl/sharedStrings.xml", "<sst><si><t>Report</t></si></sst>")))
            {
                var hits = PackageRoots.Find(package, Names);

                CollectionAssert.AreEquivalent(new[] { "OnRibbonClick", "MyUdf" }, hits.Select(h => h.Name).ToList());
            }
        }
    }
}
