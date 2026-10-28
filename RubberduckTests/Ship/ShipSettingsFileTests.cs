using System.Linq;
using NUnit.Framework;
using Rubberduck.Settings;
using Rubberduck.Ship;
using Rubberduck.SmartIndenter;

namespace RubberduckTests.Ship
{
    /// <summary>The <c>ship</c> section of rubberduck.json: read, written into a new file with a comment per setting, and updated in an existing file without touching the rest.</summary>
    [TestFixture]
    [Category("Ship")]
    public class ShipSettingsFileTests
    {
        private static ProjectConfig Read(string json) => new ProjectConfigReader(null).Parse(json, null);

        [Test]
        public void AFileWithoutAShipSectionHasNoShipSettingsAndNoErrors()
        {
            var config = Read("{ \"sync\": { \"onViolation\": \"warn\" } }");

            Assert.IsNull(config.Ship);
            Assert.IsEmpty(config.Errors);
        }

        [Test]
        public void TheShipSectionIsRead()
        {
            var config = Read(@"{
  // ship it
  ""ship"": {
    ""format"": ""xlam"",
    ""outputFolder"": ""../dist"",
    ""fileName"": ""MyAddIn"",
    ""onExisting"": ""keepBoth"",
    ""steps"": { ""comments"": false, ""notes"": true }
  }
}");

            Assert.IsEmpty(config.Errors);
            Assert.AreEqual(ShipFormat.Xlam, config.Ship.Format);
            Assert.AreEqual("../dist", config.Ship.OutputFolder);
            Assert.AreEqual("MyAddIn", config.Ship.FileName);
            Assert.AreEqual(ShipOnExisting.KeepBoth, config.Ship.OnExisting);
            Assert.IsFalse(config.Ship.IsOn(ShipSteps.Comments));
            Assert.IsTrue(config.Ship.IsOn(ShipSteps.Notes));
            Assert.IsTrue(config.Ship.IsOn(ShipSteps.Config), "A step that is not listed is on.");
        }

        [Test]
        public void KeysLeftOutKeepTheirDefaults()
        {
            var config = Read("{ \"ship\": { \"outputFolder\": \"dist\" } }");

            Assert.AreEqual(ShipFormat.Xlsm, config.Ship.Format);
            Assert.IsNull(config.Ship.FileName);
            Assert.AreEqual(ShipOnExisting.Ask, config.Ship.OnExisting);
            Assert.IsEmpty(config.Ship.Steps);
        }

        [Test]
        public void AnEmptyFolderOrNameMeansNotSet()
        {
            var config = Read("{ \"ship\": { \"outputFolder\": \"  \", \"fileName\": \"\" } }");

            Assert.IsNull(config.Ship.OutputFolder);
            Assert.IsNull(config.Ship.FileName);
        }

        [TestCase("{ \"ship\": { \"colour\": 1 } }", "Unknown ship setting 'colour'")]
        [TestCase("{ \"ship\": { \"format\": \"exe\" } }", "format")]
        [TestCase("{ \"ship\": { \"onExisting\": \"delete\" } }", "onExisting")]
        [TestCase("{ \"ship\": { \"outputFolder\": 3 } }", "outputFolder")]
        [TestCase("{ \"ship\": { \"steps\": { \"tree-shaking\": true } } }", "Unknown ship step 'tree-shaking'")]
        [TestCase("{ \"ship\": { \"steps\": { \"comments\": \"yes\" } } }", "comments")]
        [TestCase("{ \"ship\": { \"steps\": [] } }", "steps")]
        [TestCase("{ \"ship\": 5 }", "'ship' must be an object")]
        public void WhatIsWrongInTheShipSectionIsReported(string json, string expected)
        {
            var errors = Read(json).Errors;

            Assert.That(errors.Any(e => e.Contains(expected)), "Errors were: " + string.Join(" | ", errors));
        }

        [Test]
        public void AnUnknownSectionIsStillReported()
        {
            Assert.That(Read("{ \"deploy\": {} }").Errors.Any(e => e.Contains("Unknown section 'deploy'")));
        }

        [Test]
        public void ANewFileHasTheShipSectionWithACommentForEverySetting()
        {
            var text = ProjectConfigTemplate.Create(new IndenterSettings(true));
            var config = Read(text);
            var lines = text.Split(new[] { "\r\n" }, System.StringSplitOptions.None);

            Assert.IsEmpty(config.Errors, string.Join(" | ", config.Errors));
            Assert.IsNotNull(config.Ship);
            foreach (var key in new[] { "\"format\"", "\"outputFolder\"", "\"fileName\"", "\"onExisting\"", "\"steps\"" }
                         .Concat(ShipSteps.All.Select(step => "\"" + step + "\"")))
            {
                var index = System.Array.FindIndex(lines, line => line.TrimStart().StartsWith(key + ":"));
                Assert.GreaterOrEqual(index, 0, key + " is missing from the template.");
                Assert.That(lines[index - 1].TrimStart().StartsWith("//") || lines[index - 2].TrimStart().StartsWith("//"), key + " has no comment above it.");
            }
        }

        [Test]
        public void TheCommentOnTheFolderListsEveryFormItUnderstands()
        {
            var text = ProjectConfigTemplate.Create(new IndenterSettings(true));
            var start = text.IndexOf("\"ship\"");
            var comment = string.Join("\n", text.Substring(start, text.IndexOf("\"outputFolder\"") - start).Split('\n').Where(l => l.TrimStart().StartsWith("//")));

            foreach (var form in new[] { "relative", "../..", "%", "Documents", "Desktop", "Downloads", "OneDrive", "Dropbox" })
            {
                StringAssert.Contains(form, comment);
            }
        }

        [Test]
        public void WritingTheSectionReplacesOnlyTheShipSection()
        {
            var original = "{\r\n  // the indenter\r\n  \"indenter\": { \"indentSpaces\": 2 }, // trailing\r\n  \"ship\": { \"format\": \"xlsm\" },\r\n  // after\r\n  \"sync\": { \"onViolation\": \"warn\" }\r\n}\r\n";
            var settings = new ShipSettings { Format = ShipFormat.Xlam, OutputFolder = "%USERPROFILE%\\Out", FileName = "Tool", OnExisting = ShipOnExisting.Replace };
            settings.Steps[ShipSteps.Comments] = false;

            var updated = ProjectConfigWriter.SetShip(original, settings);

            StringAssert.StartsWith("{\r\n  // the indenter\r\n  \"indenter\": { \"indentSpaces\": 2 }, // trailing\r\n  \"ship\": ", updated);
            StringAssert.EndsWith(",\r\n  // after\r\n  \"sync\": { \"onViolation\": \"warn\" }\r\n}\r\n", updated);
            var config = Read(updated);
            Assert.IsEmpty(config.Errors, string.Join(" | ", config.Errors));
            Assert.AreEqual(ShipFormat.Xlam, config.Ship.Format);
            Assert.AreEqual("%USERPROFILE%\\Out", config.Ship.OutputFolder);
            Assert.AreEqual("Tool", config.Ship.FileName);
            Assert.AreEqual(ShipOnExisting.Replace, config.Ship.OnExisting);
            Assert.IsFalse(config.Ship.IsOn(ShipSteps.Comments));
            Assert.IsTrue(config.Ship.IsOn(ShipSteps.Notes));
        }

        [Test]
        public void WritingTheSectionAddsItWhenTheFileHasNone()
        {
            var original = "{\r\n  // the indenter\r\n  \"indenter\": { \"indentSpaces\": 2 } // last one\r\n}\r\n";

            var updated = ProjectConfigWriter.SetShip(original, new ShipSettings { OutputFolder = "dist" });

            StringAssert.StartsWith("{\r\n  // the indenter\r\n  \"indenter\": { \"indentSpaces\": 2 }", updated);
            StringAssert.Contains("// last one", updated);
            var config = Read(updated);
            Assert.IsEmpty(config.Errors, string.Join(" | ", config.Errors));
            Assert.AreEqual("dist", config.Ship.OutputFolder);
            Assert.AreEqual(2, config.Indenter.IndentSpaces);
        }

        [Test]
        public void WritingTheSectionIntoAnEmptyObjectWorks()
        {
            var updated = ProjectConfigWriter.SetShip("{}", new ShipSettings { OutputFolder = "dist" });

            var config = Read(updated);
            Assert.IsEmpty(config.Errors, string.Join(" | ", config.Errors));
            Assert.AreEqual("dist", config.Ship.OutputFolder);
        }

        [Test]
        public void WritingTheSectionTwiceChangesNothingTheSecondTime()
        {
            var once = ProjectConfigWriter.SetShip("{\r\n  \"sync\": {}\r\n}", new ShipSettings { OutputFolder = "dist" });
            var twice = ProjectConfigWriter.SetShip(once, new ShipSettings { OutputFolder = "dist" });

            Assert.AreEqual(once, twice);
        }

        [Test]
        public void TextThatIsNotJsonIsRefused()
        {
            Assert.Throws<Newtonsoft.Json.JsonReaderException>(() => ProjectConfigWriter.SetShip("{ nope", new ShipSettings()));
        }
    }
}
