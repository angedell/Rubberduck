using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    [TestFixture]
    public class ModuleAttributesTests
    {
        //As the VBE exports a class: the header block, the module attributes, the code, and the member attributes right under their declaration.
        private const string Exported =
            "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n" +
            "Attribute VB_Name = \"Greeter\"\r\n" +
            "Attribute VB_GlobalNameSpace = False\r\n" +
            "Attribute VB_Creatable = False\r\n" +
            "Attribute VB_PredeclaredId = False\r\n" +
            "Attribute VB_Exposed = False\r\n" +
            "Option Explicit\r\n" +
            "'@PredeclaredId\r\n" +
            "Public Function Greet(ByVal name As String) As String\r\n" +
            "    Greet = \"Hello, \" & name\r\n" +
            "End Function\r\n";

        [Test]
        public void Parse_GivesTheModuleAttributes_WithTheirTypes()
        {
            var attributes = ModuleAttributes.Parse(Exported);

            Assert.AreEqual("Greeter", (string)attributes["module"]["VB_Name"]);
            Assert.AreEqual(false, (bool)attributes["module"]["VB_PredeclaredId"]);
            Assert.AreEqual(5, ((JObject)attributes["module"]).Count);
        }

        [Test]
        public void Parse_GivesMemberAttributesUnderTheMember_AndKeepsTheirCase()
        {
            var text = Exported.Replace("    Greet = ", "Attribute Greet.VB_UserMemId = 0\r\nAttribute Greet.VB_Description = \"Says \"\"hi\"\"\"\r\n    Greet = ");

            var attributes = ModuleAttributes.Parse(text);

            Assert.AreEqual(0L, (long)attributes["members"]["Greet"]["VB_UserMemId"]);
            Assert.AreEqual("Says \"hi\"", (string)attributes["members"]["Greet"]["VB_Description"]);
        }

        [Test]
        public void Apply_ReplacesAModuleAttribute_InPlace()
        {
            var result = ModuleAttributes.Apply(Exported, JObject.Parse("{ module: { VB_PredeclaredId: true } }"));

            StringAssert.Contains("Attribute VB_PredeclaredId = True\r\n", result);
            StringAssert.DoesNotContain("VB_PredeclaredId = False", result);
            Assert.AreEqual(Exported.Split('\n').Length, result.Split('\n').Length);
        }

        [Test]
        public void Apply_AddsAModuleAttribute_AfterTheLastOne()
        {
            var result = ModuleAttributes.Apply(Exported, JObject.Parse("{ module: { VB_Description: 'Greets people' } }"));

            Assert.IsTrue(result.Contains("Attribute VB_Exposed = False\r\nAttribute VB_Description = \"Greets people\"\r\nOption Explicit"));
        }

        [Test]
        public void Apply_AddsMemberAttributes_RightUnderTheDeclaration()
        {
            var result = ModuleAttributes.Apply(Exported, JObject.Parse("{ members: { Greet: { VB_UserMemId: 0, VB_Description: 'Greets' } } }"));

            StringAssert.Contains("Public Function Greet(ByVal name As String) As String\r\nAttribute Greet.VB_UserMemId = 0\r\nAttribute Greet.VB_Description = \"Greets\"\r\n    Greet =", result);
        }

        [Test]
        public void Apply_PutsMemberAttributesBelowAContinuedDeclaration()
        {
            var text = Exported.Replace("(ByVal name As String) As String", "(ByVal name As String, _\r\n    ByVal greeting As String) As String");

            var result = ModuleAttributes.Apply(text, JObject.Parse("{ members: { Greet: { VB_UserMemId: 0 } } }"));

            StringAssert.Contains("ByVal greeting As String) As String\r\nAttribute Greet.VB_UserMemId = 0\r\n", result);
        }

        [Test]
        public void ParseAfterApply_GivesWhatWasSet()
        {
            var changes = JObject.Parse("{ module: { VB_PredeclaredId: true }, members: { Greet: { VB_UserMemId: 0 } } }");

            var attributes = ModuleAttributes.Parse(ModuleAttributes.Apply(Exported, changes));

            Assert.AreEqual(true, (bool)attributes["module"]["VB_PredeclaredId"]);
            Assert.AreEqual(0L, (long)attributes["members"]["Greet"]["VB_UserMemId"]);
        }

        [Test]
        public void Apply_RefusesAMemberThatIsNotThere()
        {
            var failure = Assert.Throws<ApiException>(() => ModuleAttributes.Apply(Exported, JObject.Parse("{ members: { Nope: { VB_UserMemId: 0 } } }")));

            Assert.AreEqual(400, failure.Status);
        }

        [TestCase("{ module: { Description: 'x' } }")]
        [TestCase("{ module: { VB_Description: [1] } }")]
        public void Apply_RefusesWhatIsNotAnAttribute(string changes)
            => Assert.Throws<ApiException>(() => ModuleAttributes.Apply(Exported, JObject.Parse(changes)));
    }
}
