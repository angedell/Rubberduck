using NUnit.Framework;
using Rubberduck.Ship;
using static RubberduckTests.Ship.ShipTestParser;

namespace RubberduckTests.Ship
{
    /// <summary>A module that says Option Explicit twice does not compile, and the VBE tells the user in a dialog.</summary>
    [TestFixture]
    [Category("Ship")]
    public class ShipCodeTextTests
    {
        [Test]
        public void ARepeatedOptionExplicitIsTakenOut()
        {
            var code = "Option Explicit\r\nOption Explicit\r\nPublic Sub Foo()\r\nEnd Sub";

            Assert.AreEqual("Option Explicit\r\nPublic Sub Foo()\r\nEnd Sub", ShipCodeText.WithoutRepeatedOptions(code));
        }

        [Test]
        public void DifferentOptionsAndASingleOneStay()
        {
            var code = "Option Compare Text\r\nOption Explicit\r\nOption Private Module\r\nPublic Sub Foo()\r\nEnd Sub";

            Assert.AreEqual(code, ShipCodeText.WithoutRepeatedOptions(code));
        }

        [Test]
        public void ACaseOrSpacingDifferenceIsTheSameStatement()
        {
            Assert.AreEqual("Option Explicit\r\n", ShipCodeText.WithoutRepeatedOptions("Option Explicit\r\noption   explicit\r\n"));
        }

        [Test]
        public void StrippedCodeThatStartsWithOptionExplicitHasExactlyOne()
        {
            // What the writer puts into a document module: the code of the stripped export. Added to a module that already says Option Explicit it must still end up with one.
            var stripped = new ShipModuleText(ShipTestParser.Create()).CodeOnly("ThisWorkbook", "project-1",
                Lines("VERSION 1.0 CLASS", "BEGIN", "  MultiUse = -1", "END", "Attribute VB_Name = \"ThisWorkbook\"", "Option Explicit", "' comment", "Public x As Long"));
            var inModule = "Option Explicit\r\n" + stripped.Text;

            Assert.AreEqual(Lines("Option Explicit", "' comment", "Public x As Long").Replace("' comment\r\n", string.Empty) + string.Empty,
                ShipCodeText.WithoutRepeatedOptions(inModule).Replace("' comment\r\n", string.Empty));
        }
    }
}
