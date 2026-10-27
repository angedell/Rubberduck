using System.Linq;
using NUnit.Framework;
using Rubberduck.SmartIndenter;

namespace RubberduckTests.SmartIndenter
{
    [TestFixture]
    public class VerticalSpaceTests
    {
        private static string[] Indent(string[] code, IndenterSettings settings) =>
            new Indenter(null, () => settings).Indent(code, settings).ToArray();

        private static readonly string[] Messy =
        {
            "Option Explicit",
            "",
            "",
            "",
            "Public Sub A()",
            "",
            "Debug.Print 1",
            "",
            "",
            "",
            "Debug.Print 2",
            "",
            "End Sub"
        };

        [Test]
        [Category("Indenter")]
        public void ByDefaultEmptyLinesAreLeftAlone()
        {
            var actual = Indent(Messy, new IndenterSettings());

            CollectionAssert.AreEqual(
                new[] { "Option Explicit", "", "Public Sub A()", "", "    Debug.Print 1", "", "", "", "    Debug.Print 2", "", "End Sub" },
                actual);
        }

        [Test]
        [Category("Indenter")]
        public void MaximumConsecutiveEmptyLinesCollapsesRuns()
        {
            var actual = Indent(Messy, new IndenterSettings { MaximumConsecutiveEmptyLines = 1 });

            CollectionAssert.AreEqual(
                new[] { "Option Explicit", "", "Public Sub A()", "", "    Debug.Print 1", "", "    Debug.Print 2", "", "End Sub" },
                actual);
        }

        [Test]
        [Category("Indenter")]
        public void MaximumConsecutiveEmptyLinesWinsOverLinesBetweenProcedures()
        {
            var code = new[] { "Public Sub A()", "End Sub", "Public Sub B()", "End Sub" };

            var actual = Indent(code, new IndenterSettings { LinesBetweenProcedures = 2, MaximumConsecutiveEmptyLines = 1 });

            CollectionAssert.AreEqual(new[] { "Public Sub A()", "End Sub", "", "Public Sub B()", "End Sub" }, actual);
        }

        [Test]
        [Category("Indenter")]
        public void MaximumConsecutiveEmptyLinesOfTwoKeepsTwoBetweenProcedures()
        {
            var code = new[] { "Public Sub A()", "End Sub", "Public Sub B()", "End Sub" };

            var actual = Indent(code, new IndenterSettings { LinesBetweenProcedures = 2, MaximumConsecutiveEmptyLines = 2 });

            CollectionAssert.AreEqual(new[] { "Public Sub A()", "End Sub", "", "", "Public Sub B()", "End Sub" }, actual);
        }

        [Test]
        [Category("Indenter")]
        public void RemoveEmptyLinesAtProcedureEdgesKeepsTheOnesInside()
        {
            var actual = Indent(Messy, new IndenterSettings { RemoveEmptyLinesAtProcedureEdges = true });

            CollectionAssert.AreEqual(
                new[] { "Option Explicit", "", "Public Sub A()", "    Debug.Print 1", "", "", "", "    Debug.Print 2", "End Sub" },
                actual);
        }

        [Test]
        [Category("Indenter")]
        public void BothSettingsTogether()
        {
            var actual = Indent(Messy, new IndenterSettings { RemoveEmptyLinesAtProcedureEdges = true, MaximumConsecutiveEmptyLines = 1 });

            CollectionAssert.AreEqual(
                new[] { "Option Explicit", "", "Public Sub A()", "    Debug.Print 1", "", "    Debug.Print 2", "End Sub" },
                actual);
        }

        [Test]
        [Category("Indenter")]
        public void RemoveLeadingEmptyLinesStripsTheTopOfTheModule()
        {
            var code = new[] { "", "", "Option Explicit", "Public Sub A()", "End Sub" };

            var actual = Indent(code, new IndenterSettings { RemoveLeadingEmptyLines = true });

            Assert.AreEqual("Option Explicit", actual.First());
        }

        [Test]
        [Category("Indenter")]
        public void LeadingEmptyLinesStayByDefault()
        {
            var code = new[] { "", "Option Explicit" };

            Assert.AreEqual("", Indent(code, new IndenterSettings()).First());
        }

        [Test]
        [Category("Indenter")]
        public void LimitsAreClamped()
        {
            Assert.AreEqual(IndenterSettings.MaximumEmptyLinesLimit, new IndenterSettings { MaximumConsecutiveEmptyLines = 99 }.MaximumConsecutiveEmptyLines);
            Assert.AreEqual(0, new IndenterSettings { MaximumConsecutiveEmptyLines = -3 }.MaximumConsecutiveEmptyLines);
        }

        [Test]
        [Category("Indenter")]
        public void NewSettingsCountForEquality()
        {
            Assert.AreNotEqual(new IndenterSettings(), new IndenterSettings { MaximumConsecutiveEmptyLines = 1 });
            Assert.AreNotEqual(new IndenterSettings(), new IndenterSettings { RemoveEmptyLinesAtProcedureEdges = true });
            Assert.AreNotEqual(new IndenterSettings(), new IndenterSettings { RemoveLeadingEmptyLines = true });
            Assert.AreEqual(new IndenterSettings(), new IndenterSettings());
        }
    }
}
