using NUnit.Framework;
using Rubberduck.Ship;
using Rubberduck.VBEditor.SafeComWrappers;
using static RubberduckTests.Ship.ShipTestParser;

namespace RubberduckTests.Ship
{
    /// <summary>ADR 0003, section 4: functional attributes ship, the rest does not.</summary>
    [TestFixture]
    [Category("Ship")]
    public class ShipAttributeTests
    {
        private static ShipTextResult Strip(string text, ComponentType type = ComponentType.StandardModule)
            => new ShipModuleText(ShipTestParser.Create()).StripDescriptiveAttributes("Module1", "project-1", text, type);

        [Test]
        public void TheFunctionalModuleAttributesStayAndEveryExtKeyGoes()
        {
            var result = Strip(Lines(
                "VERSION 1.0 CLASS",
                "BEGIN",
                "  MultiUse = -1  'True",
                "END",
                "Attribute VB_Name = \"Thing\"",
                "Attribute VB_GlobalNameSpace = False",
                "Attribute VB_Creatable = False",
                "Attribute VB_PredeclaredId = True",
                "Attribute VB_Exposed = True",
                "Attribute VB_Description = \"A thing\"",
                "Attribute VB_Ext_KEY = \"Rubberduck\" ,\"Predeclared\"",
                "Attribute VB_Ext_KEY = \"Annotation\" ,\"Folder\"",
                "Option Explicit",
                "Public x As Long"), ComponentType.ClassModule);

            Assert.IsTrue(result.Parsed, result.Problem);
            Assert.AreEqual(Lines(
                "VERSION 1.0 CLASS",
                "BEGIN",
                "  MultiUse = -1  'True",
                "END",
                "Attribute VB_Name = \"Thing\"",
                "Attribute VB_GlobalNameSpace = False",
                "Attribute VB_Creatable = False",
                "Attribute VB_PredeclaredId = True",
                "Attribute VB_Exposed = True",
                "Option Explicit",
                "Public x As Long"), result.Text);
            Assert.AreEqual(3, result.Removed);
        }

        [Test]
        public void DescriptionAndInvokeDataStayOnPublicProceduresOfAStandardModule()
        {
            var code = Lines(
                "Attribute VB_Name = \"Module1\"",
                "Public Sub Run()",
                "Attribute Run.VB_Description = \"Runs it\"",
                "Attribute Run.VB_ProcData.VB_Invoke_Func = \" \\n14\"",
                "End Sub",
                "Function Total() As Long",
                "Attribute Total.VB_Description = \"No modifier means public\"",
                "End Function");
            var result = Strip(code);

            Assert.AreEqual(code, result.Text);
            Assert.AreEqual(0, result.Removed);
        }

        [Test]
        public void DescriptionsOfPrivateProceduresGo()
        {
            var result = Strip(Lines(
                "Attribute VB_Name = \"Module1\"",
                "Private Sub Hidden()",
                "Attribute Hidden.VB_Description = \"private\"",
                "Attribute Hidden.VB_ProcData.VB_Invoke_Func = \" \\n14\"",
                "End Sub"));

            Assert.AreEqual(Lines("Attribute VB_Name = \"Module1\"", "Private Sub Hidden()", "End Sub"), result.Text);
        }

        [Test]
        public void NothingIsKeptWhenTheModuleIsPrivate()
        {
            var result = Strip(Lines(
                "Attribute VB_Name = \"Module1\"",
                "Option Private Module",
                "Public Sub Run()",
                "Attribute Run.VB_Description = \"Runs it\"",
                "End Sub"));

            Assert.AreEqual(Lines("Attribute VB_Name = \"Module1\"", "Option Private Module", "Public Sub Run()", "End Sub"), result.Text);
        }

        [Test]
        public void DescriptionsInAClassModuleGoButTheDefaultMemberStays()
        {
            var result = Strip(Lines(
                "VERSION 1.0 CLASS",
                "BEGIN",
                "  MultiUse = -1  'True",
                "END",
                "Attribute VB_Name = \"Bag\"",
                "Public Function Item(ByVal i As Long) As Variant",
                "Attribute Item.VB_Description = \"The item\"",
                "Attribute Item.VB_UserMemId = 0",
                "End Function"), ComponentType.ClassModule);

            Assert.AreEqual(Lines(
                "VERSION 1.0 CLASS",
                "BEGIN",
                "  MultiUse = -1  'True",
                "END",
                "Attribute VB_Name = \"Bag\"",
                "Public Function Item(ByVal i As Long) As Variant",
                "Attribute Item.VB_UserMemId = 0",
                "End Function"), result.Text);
        }

        [Test]
        public void VariableDescriptionsGo_ButMemberFlagsAndVariableMemberIdsStay()
        {
            var result = Strip(Lines(
                "Attribute VB_Name = \"Module1\"",
                "Public Counter As Long",
                "Attribute Counter.VB_VarDescription = \"counts\"",
                "Attribute Counter.VB_VarUserMemId = 0",
                "Public Sub Run()",
                "Attribute Run.VB_MemberFlags = \"40\"",
                "End Sub"));

            Assert.AreEqual(Lines(
                "Attribute VB_Name = \"Module1\"",
                "Public Counter As Long",
                "Attribute Counter.VB_VarUserMemId = 0",
                "Public Sub Run()",
                "Attribute Run.VB_MemberFlags = \"40\"",
                "End Sub"), result.Text);
        }

        [Test]
        public void AModuleThatDoesNotParseIsLeftAlone()
        {
            var code = Lines("Attribute VB_Name = \"Module1\"", "Attribute VB_Ext_KEY = \"a\" ,\"b\"", "Public Sub Run(");
            var result = Strip(code);

            Assert.IsFalse(result.Parsed);
            Assert.AreEqual(code, result.Text);
        }

        [Test]
        public void CodeOnlyDropsTheHeaderTheDesignerBlockAndEveryAttribute()
        {
            var result = new ShipModuleText(ShipTestParser.Create()).CodeOnly("Thing", "project-1", Lines(
                "VERSION 1.0 CLASS",
                "BEGIN",
                "  MultiUse = -1  'True",
                "END",
                "Attribute VB_Name = \"Thing\"",
                "Attribute VB_Exposed = True",
                "Option Explicit",
                "Public Sub Run()",
                "Attribute Run.VB_Description = \"Runs it\"",
                "    Debug.Print 1",
                "End Sub"));

            Assert.IsTrue(result.Parsed, result.Problem);
            Assert.AreEqual(Lines("Option Explicit", "Public Sub Run()", "    Debug.Print 1", "End Sub"), result.Text);
        }
    }
}
