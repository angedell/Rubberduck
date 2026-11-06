using System.Linq;
using NUnit.Framework;
using Rubberduck.Settings;
using Rubberduck.Ship;
using Rubberduck.SmartIndenter;

namespace RubberduckTests.Ship
{
    /// <summary><c>ship.steps.treeShaking</c> and <c>ship.treeShaking</c> in rubberduck.json: defaults, reading, errors, the template and the round trip.</summary>
    [TestFixture]
    [Category("Ship")]
    public class ShipTreeShakingSettingsTests
    {
        private static ProjectConfig Read(string json) => new ProjectConfigReader(null).Parse(json, null);

        [Test]
        public void TreeShakingIsOffUntilTheFileSaysOtherwise()
        {
            Assert.IsFalse(new ShipSettings().IsOn(ShipSteps.TreeShaking));
            Assert.IsFalse(Read("{ \"ship\": { } }").Ship.IsOn(ShipSteps.TreeShaking));
            Assert.IsTrue(Read("{ \"ship\": { \"steps\": { \"treeShaking\": true } } }").Ship.IsOn(ShipSteps.TreeShaking));
            Assert.IsTrue(new ShipSettings().IsOn(ShipSteps.Comments), "The other steps are still on by default.");
        }

        [Test]
        public void TheModeAndTheKeepListAreRead()
        {
            var config = Read(@"{ ""ship"": { ""treeShaking"": { ""mode"": ""EntryPointsOnly"", ""keep"": [ ""Module1.Run"", "" OnRibbonClick "" ] } } }");

            Assert.IsEmpty(config.Errors, string.Join(" | ", config.Errors));
            Assert.AreEqual(ShipTreeShakingMode.EntryPointsOnly, config.Ship.TreeShaking.Mode);
            CollectionAssert.AreEqual(new[] { "Module1.Run", "OnRibbonClick" }, config.Ship.TreeShaking.Keep);
        }

        [Test]
        public void WithoutTheObjectTheModeIsConservativeAndTheKeepListEmpty()
        {
            var config = Read("{ \"ship\": { \"format\": \"xlam\" } }");

            Assert.AreEqual(ShipTreeShakingMode.Conservative, config.Ship.TreeShaking.Mode);
            Assert.IsEmpty(config.Ship.TreeShaking.Keep);
        }

        [TestCase("{ \"ship\": { \"treeShaking\": { \"mode\": \"aggressive\" } } }", "treeShaking.mode")]
        [TestCase("{ \"ship\": { \"treeShaking\": { \"mode\": 1 } } }", "treeShaking.mode")]
        [TestCase("{ \"ship\": { \"treeShaking\": { \"keep\": \"Module1.Run\" } } }", "treeShaking.keep")]
        [TestCase("{ \"ship\": { \"treeShaking\": { \"keep\": [ 3 ] } } }", "treeShaking.keep")]
        [TestCase("{ \"ship\": { \"treeShaking\": { \"colour\": 1 } } }", "Unknown ship setting 'treeShaking.colour'")]
        [TestCase("{ \"ship\": { \"treeShaking\": [] } }", "'treeShaking' must be an object")]
        [TestCase("{ \"ship\": { \"steps\": { \"treeShaking\": \"yes\" } } }", "treeShaking")]
        public void WhatIsWrongWithTheTreeShakingSettingsIsReported(string json, string expected)
        {
            var errors = Read(json).Errors;

            Assert.That(errors.Any(e => e.Contains(expected)), "Errors were: " + string.Join(" | ", errors));
        }

        [Test]
        public void KeepEntriesThatBreakTheLimitsAreReportedAndLeftOut()
        {
            var tooLong = new string('x', ShipKeepList.MaxLength + 1);
            var config = Read("{ \"ship\": { \"treeShaking\": { \"keep\": [ \"Good\", \"\", \"   \", \"" + tooLong + "\", \"Bad\\u0007Name\", \"Fine\" ] } } }");

            Assert.That(config.Errors.Any(e => e.Contains("empty")), string.Join(" | ", config.Errors));
            Assert.That(config.Errors.Any(e => e.Contains("longer than")));
            Assert.That(config.Errors.Any(e => e.Contains("control character")));
            CollectionAssert.AreEqual(new[] { "Good", "Fine" }, config.Ship.TreeShaking.Keep);
        }

        [Test]
        public void MoreThanTheLimitOfKeepEntriesIsReportedAndTheRestIgnored()
        {
            var many = string.Join(", ", Enumerable.Range(0, ShipKeepList.MaxEntries + 3).Select(i => "\"N" + i + "\""));
            var config = Read("{ \"ship\": { \"treeShaking\": { \"keep\": [ " + many + " ] } } }");

            Assert.That(config.Errors.Any(e => e.Contains("at most")), string.Join(" | ", config.Errors));
            Assert.AreEqual(ShipKeepList.MaxEntries, config.Ship.TreeShaking.Keep.Count);
            Assert.AreEqual("N0", config.Ship.TreeShaking.Keep[0]);
        }

        [Test]
        public void TheSharedValidatorKeepsTheGoodEntriesInOrder()
        {
            var errors = new System.Collections.Generic.List<string>();

            var valid = ShipKeepList.Validate(new[] { " A ", "", "B" }, errors);

            CollectionAssert.AreEqual(new[] { "A", "B" }, valid);
            Assert.AreEqual(1, errors.Count);
        }

        [Test]
        public void ANewFileHasTheStepOffAndTheObjectWithACommentForEverySetting()
        {
            var text = ProjectConfigTemplate.Create(new IndenterSettings(true));
            var config = Read(text);
            var lines = text.Split(new[] { "\r\n" }, System.StringSplitOptions.None);

            Assert.IsEmpty(config.Errors, string.Join(" | ", config.Errors));
            Assert.IsFalse(config.Ship.IsOn(ShipSteps.TreeShaking));
            Assert.AreEqual(ShipTreeShakingMode.Conservative, config.Ship.TreeShaking.Mode);
            foreach (var key in new[] { "\"treeShaking\": {", "\"mode\":", "\"keep\":" })
            {
                var index = System.Array.FindIndex(lines, line => line.TrimStart().StartsWith(key));
                Assert.GreaterOrEqual(index, 0, key + " is missing from the template.");
                Assert.That(lines[index - 1].TrimStart().StartsWith("//"), key + " has no comment above it.");
            }

            var step = System.Array.FindIndex(lines, line => line.TrimStart().StartsWith("\"treeShaking\": false"));
            Assert.GreaterOrEqual(step, 0, "The step is listed, off.");
            Assert.That(lines[step - 1].TrimStart().StartsWith("//"));
        }

        [Test]
        public void WhatIsWrittenIsReadBackAndWritingTwiceChangesNothing()
        {
            var settings = new ShipSettings { OutputFolder = "dist" };
            settings.Steps[ShipSteps.TreeShaking] = true;
            settings.TreeShaking.Mode = ShipTreeShakingMode.EntryPointsOnly;
            settings.TreeShaking.Keep.AddRange(new[] { "Module1.Run", "OnRibbonClick", "Quote\"d" });

            var once = ProjectConfigWriter.SetShip("{\r\n  \"sync\": {}\r\n}", settings);
            var back = Read(once);
            var twice = ProjectConfigWriter.SetShip(once, back.Ship);

            Assert.IsEmpty(back.Errors, string.Join(" | ", back.Errors));
            Assert.IsTrue(back.Ship.IsOn(ShipSteps.TreeShaking));
            Assert.AreEqual(ShipTreeShakingMode.EntryPointsOnly, back.Ship.TreeShaking.Mode);
            CollectionAssert.AreEqual(settings.TreeShaking.Keep, back.Ship.TreeShaking.Keep);
            Assert.AreEqual(once, twice);
        }

        [Test]
        public void ReplacingAnExistingSectionKeepsTheRestOfTheFile()
        {
            var original = "{\r\n  // before\r\n  \"ship\": { \"treeShaking\": { \"mode\": \"entryPointsOnly\", \"keep\": [\"A\"] } },\r\n  // after\r\n  \"sync\": {}\r\n}\r\n";
            var settings = Read(original).Ship;
            settings.TreeShaking.Keep.Add("B");

            var updated = ProjectConfigWriter.SetShip(original, settings);

            StringAssert.StartsWith("{\r\n  // before\r\n  \"ship\": ", updated);
            StringAssert.EndsWith(",\r\n  // after\r\n  \"sync\": {}\r\n}\r\n", updated);
            CollectionAssert.AreEqual(new[] { "A", "B" }, Read(updated).Ship.TreeShaking.Keep);
        }
    }
}
