using System.Linq;
using NUnit.Framework;
using Rubberduck.Refactorings.ExtractUdtClass;

namespace RubberduckTests.Refactoring.ExtractUdtClass
{
    [TestFixture]
    public class ExtractUdtClassCodeBuilderTests
    {
        private static readonly UdtMemberInfo[] PersonMembers =
        {
            new UdtMemberInfo("Name", "String", false, false, false, "Name As String"),
            new UdtMemberInfo("Age", "Long", false, false, false, "Age As Long"),
        };

        private static string Lf(string text) => text.Replace("\r\n", "\n");

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void PlainClassHasPropertiesAndSelf()
        {
            const string expected =
@"' Usage:
'  Dim item As Person
'  Set item = New Person
'  item.Name = ""text""
'  item.Age = 0
'  Debug.Print item.Name
' Self returns the object itself, which suits With blocks and chaining.
Option Explicit

Private Type TPerson
    Name As String
    Age As Long
End Type

Private This As TPerson

Public Property Get Self() As Person
    Set Self = Me
End Property

Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Let Name(ByVal Value As String)
    This.Name = Value
End Property

Public Property Get Age() As Long
    Age = This.Age
End Property

Public Property Let Age(ByVal Value As Long)
    This.Age = Value
End Property
";

            var actual = ExtractUdtClassCodeBuilder.ClassBody("Person", "TPerson", PersonMembers, ExtractUdtClassStyle.Plain);

            Assert.AreEqual(Lf(expected), Lf(actual));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void FluentClassIsPredeclaredWithCreateAndWithMethods()
        {
            var actual = Lf(ExtractUdtClassCodeBuilder.ClassBody("Person", "TPerson", PersonMembers, ExtractUdtClassStyle.Fluent));

            StringAssert.StartsWith("'@PredeclaredId\n' Usage:\n", actual);
            StringAssert.Contains("Self returns the object itself, which suits With blocks and chaining.\nOption Explicit\n", actual);
            StringAssert.Contains("Public Function Create() As Person\n    Set Create = New Person\nEnd Function\n", actual);
            StringAssert.Contains("Public Function WithName(ByVal Value As String) As Person\n    This.Name = Value\n    Set WithName = Me\nEnd Function\n", actual);
            StringAssert.Contains("Public Function WithAge(ByVal Value As Long) As Person\n", actual);
            StringAssert.Contains("Public Property Get Self() As Person\n", actual);
            StringAssert.Contains("Public Property Let Name(ByVal Value As String)\n", actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void PlainClassHasNoCreateOrWithMethods()
        {
            var actual = ExtractUdtClassCodeBuilder.ClassBody("Person", "TPerson", PersonMembers, ExtractUdtClassStyle.Plain);

            StringAssert.DoesNotContain("Create", actual);
            StringAssert.DoesNotContain("WithName", actual);
            StringAssert.DoesNotContain("WithAge", actual);
            StringAssert.DoesNotContain("PredeclaredId", actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void ObjectMembersUsePropertySetAndSet()
        {
            var members = new[] { new UdtMemberInfo("Owner", "Collection", false, true, false, "Owner As Collection") };

            var actual = Lf(ExtractUdtClassCodeBuilder.ClassBody("Bag", "TBag", members, ExtractUdtClassStyle.Fluent));

            StringAssert.Contains("Public Property Get Owner() As Collection\n    Set Owner = This.Owner\n", actual);
            StringAssert.Contains("Public Property Set Owner(ByVal Value As Collection)\n    Set This.Owner = Value\n", actual);
            StringAssert.Contains("Public Function WithOwner(ByVal Value As Collection) As Bag\n    Set This.Owner = Value\n    Set WithOwner = Me\n", actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void ArrayMembersGetOnlyAGetter_AndKeepTheirDeclaredBounds()
        {
            var members = new[] { new UdtMemberInfo("Scores", "Long", true, false, false, "Scores(1 To 3) As Long") };

            var actual = Lf(ExtractUdtClassCodeBuilder.ClassBody("Result", "TResult", members, ExtractUdtClassStyle.Fluent));

            StringAssert.Contains("    Scores(1 To 3) As Long\n", actual);
            StringAssert.Contains("Public Property Get Scores() As Long()\n    Scores = This.Scores\n", actual);
            StringAssert.DoesNotContain("Property Let Scores", actual);
            StringAssert.DoesNotContain("WithScores", actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void MemberNamedValueDoesNotClashWithTheParameter()
        {
            var members = new[] { new UdtMemberInfo("Value", "Long", false, false, false, "Value As Long") };

            var actual = Lf(ExtractUdtClassCodeBuilder.ClassBody("Cell", "TCell", members, ExtractUdtClassStyle.Plain));

            StringAssert.Contains("Public Property Let Value(ByVal NewValue As Long)\n    This.Value = NewValue\n", actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void NestedTypesAreNotExposedButLeftAsAComment()
        {
            var members = new[] { new UdtMemberInfo("Home", "TAddress", false, false, true, "Home As TAddress") };

            var actual = Lf(ExtractUdtClassCodeBuilder.ClassBody("Person", "TPerson", members, ExtractUdtClassStyle.Fluent));

            StringAssert.Contains("' Home is a user defined type;", actual);
            StringAssert.DoesNotContain("Property Get Home", actual);
            StringAssert.DoesNotContain("WithHome", actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void FolderAnnotationIsKept()
        {
            var actual = Lf(ExtractUdtClassCodeBuilder.ClassBody("Person", "TPerson", PersonMembers, ExtractUdtClassStyle.Plain, "'@Folder(\"Model\")"));

            StringAssert.StartsWith("'@Folder(\"Model\")\n' Usage:\n", actual);
        }

        private static UdtMemberInfo Member(string name) => new UdtMemberInfo(name, "Long", false, false, false, name + " As Long");

        [TestCase("Self", ExtractUdtClassStyle.Plain)]
        [TestCase("self", ExtractUdtClassStyle.Fluent)]
        [TestCase("This", ExtractUdtClassStyle.Plain)]
        [TestCase("Create", ExtractUdtClassStyle.Fluent)]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void MembersNamedLikeWhatTheClassAddsAreClashes(string name, ExtractUdtClassStyle style)
        {
            var clashes = ExtractUdtClassCodeBuilder.Clashes(new[] { Member("Age"), Member(name) }, style);

            Assert.AreEqual(1, clashes.Count);
            StringAssert.Contains($"'{name}'", clashes[0]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void CreateOnlyClashesInTheFluentStyle()
        {
            CollectionAssert.IsEmpty(ExtractUdtClassCodeBuilder.Clashes(new[] { Member("Create") }, ExtractUdtClassStyle.Plain));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void AMemberCalledWithXClashesWithTheWithXMethodOfX_OnlyInTheFluentStyle()
        {
            var members = new[] { Member("Name"), Member("WithName") };

            var clashes = ExtractUdtClassCodeBuilder.Clashes(members, ExtractUdtClassStyle.Fluent);

            Assert.AreEqual(1, clashes.Count);
            StringAssert.Contains("'WithName' clashes with the WithName method that 'Name' gets", clashes[0]);
            CollectionAssert.IsEmpty(ExtractUdtClassCodeBuilder.Clashes(members, ExtractUdtClassStyle.Plain));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void AnArrayHasNoWithMethod_SoNothingClashesWithIt()
        {
            var members = new[] { new UdtMemberInfo("Name", "Long", true, false, false, "Name(1 To 2) As Long"), Member("WithName") };

            CollectionAssert.IsEmpty(ExtractUdtClassCodeBuilder.Clashes(members, ExtractUdtClassStyle.Fluent));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void OrdinaryNamesDoNotClash()
        {
            CollectionAssert.IsEmpty(ExtractUdtClassCodeBuilder.Clashes(PersonMembers, ExtractUdtClassStyle.Fluent));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void TheModuleHeaderUsesCarriageReturnLineFeedsOnly()
        {
            var header = ExtractUdtClassCodeBuilder.ModuleHeader("Person");

            StringAssert.StartsWith("VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Person\"\r\n", header);
            StringAssert.EndsWith("Attribute VB_Exposed = False", header);
            Assert.AreEqual(8, header.Count(c => c == '\n'));
            Assert.AreEqual(8, header.Count(c => c == '\r'));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void TheWithMethodStartsWithACapital_WhateverTheMemberLooksLike()
        {
            var members = new[] { new UdtMemberInfo("name", "String", false, false, false, "name As String") };

            var actual = Lf(ExtractUdtClassCodeBuilder.ClassBody("Person", "TPerson", members, ExtractUdtClassStyle.Fluent));

            StringAssert.Contains("Public Function WithName(ByVal Value As String) As Person\n    This.name = Value\n    Set WithName = Me\n", actual);
        }

        private static UdtMemberInfo Required(string name, string type = "String", bool isObject = false)
            => new UdtMemberInfo(name, type, false, isObject, false, name + " As " + type) { IsRequired = true };

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void RequiredMembersBecomeTheArgumentsOfCreate_InTheOrderOfTheType()
        {
            var members = new[] { Required("Name"), new UdtMemberInfo("Age", "Long", false, false, false, "Age As Long"), Required("Born", "Date") };

            var actual = Lf(ExtractUdtClassCodeBuilder.ClassBody("Person", "TPerson", members, ExtractUdtClassStyle.Fluent));

            StringAssert.Contains(
                "Public Function Create(ByVal Name As String, ByVal Born As Date) As Person\n"
                + "    Dim NewInstance As Person\n"
                + "    Set NewInstance = New Person\n"
                + "    NewInstance.Name = Name\n"
                + "    NewInstance.Born = Born\n"
                + "    Set Create = NewInstance\n"
                + "End Function\n", actual);
            //The member can still be changed afterwards.
            StringAssert.Contains("Public Function WithName(", actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void ARequiredObjectIsSet_NotLet()
        {
            var members = new[] { Required("Owner", "Collection", isObject: true) };

            var actual = Lf(ExtractUdtClassCodeBuilder.ClassBody("Bag", "TBag", members, ExtractUdtClassStyle.Fluent));

            StringAssert.Contains("Public Function Create(ByVal Owner As Collection) As Bag\n", actual);
            StringAssert.Contains("    Set NewInstance.Owner = Owner\n", actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void TheNewObjectDoesNotTakeTheNameOfAnArgument()
        {
            var members = new[] { Required("NewInstance") };

            var actual = Lf(ExtractUdtClassCodeBuilder.ClassBody("Thing", "TThing", members, ExtractUdtClassStyle.Fluent));

            StringAssert.Contains("Public Function Create(ByVal NewInstance As String) As Thing\n    Dim NewInstance_ As Thing\n", actual);
            StringAssert.Contains("    NewInstance_.NewInstance = NewInstance\n", actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void RequiredMembersMeanNothingWithoutCreate_AndArraysCannotBeRequired()
        {
            var plain = Lf(ExtractUdtClassCodeBuilder.ClassBody("Person", "TPerson", new[] { Required("Name") }, ExtractUdtClassStyle.Plain));
            StringAssert.DoesNotContain("Create", plain);

            var array = new UdtMemberInfo("Scores", "Long", true, false, false, "Scores(1 To 3) As Long") { IsRequired = true };
            var fluent = Lf(ExtractUdtClassCodeBuilder.ClassBody("Result", "TResult", new[] { array }, ExtractUdtClassStyle.Fluent));
            StringAssert.Contains("Public Function Create() As Result\n    Set Create = New Result\n", fluent);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void TheFluentUsageShowsCreateWithTheRequiredArguments_AndChainsTheOthers()
        {
            var members = new[] { Required("Name"), new UdtMemberInfo("Age", "Long", false, false, false, "Age As Long"), new UdtMemberInfo("Born", "Date", false, false, false, "Born As Date") };

            var actual = Lf(ExtractUdtClassCodeBuilder.UsageComment("Person", members, ExtractUdtClassStyle.Fluent));

            Assert.AreEqual(
                "' Usage:\n"
                + "'  Dim item As Person\n"
                + "'  Set item = Person.Create(\"text\") _\n"
                + "'      .WithAge(0) _\n"
                + "'      .WithBorn(#1/1/2000#)\n"
                + "'  Debug.Print item.Name\n"
                + "' Self returns the object itself, which suits With blocks and chaining.\n", actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void TheFluentUsageIsOneLineWhenThereIsNothingToChain()
        {
            var actual = Lf(ExtractUdtClassCodeBuilder.UsageComment("Person", new[] { Required("Name") }, ExtractUdtClassStyle.Fluent));

            StringAssert.Contains("'  Set item = Person.Create(\"text\")\n'  Debug.Print", actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void TheUsageSetsObjectsWithSet_AndSkipsArraysAndNestedTypes()
        {
            var members = new[]
            {
                new UdtMemberInfo("Owner", "Collection", false, true, false, "Owner As Collection"),
                new UdtMemberInfo("Scores", "Long", true, false, false, "Scores(1 To 3) As Long"),
                new UdtMemberInfo("Home", "TAddress", false, false, true, "Home As TAddress"),
            };

            var actual = Lf(ExtractUdtClassCodeBuilder.UsageComment("Bag", members, ExtractUdtClassStyle.Plain));

            StringAssert.Contains("'  Set item.Owner = Nothing\n", actual);
            StringAssert.DoesNotContain("Scores", actual);
            StringAssert.DoesNotContain("Home", actual);
            StringAssert.DoesNotContain("Debug.Print", actual);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void EveryUsageLineIsAComment()
        {
            foreach (var style in new[] { ExtractUdtClassStyle.Plain, ExtractUdtClassStyle.Fluent })
            {
                var lines = ExtractUdtClassCodeBuilder.UsageComment("Person", PersonMembers, style)
                    .Split(new[] { System.Environment.NewLine }, System.StringSplitOptions.RemoveEmptyEntries);

                Assert.IsTrue(lines.All(line => line.StartsWith("'")));
            }
        }

        [TestCase("TPerson", "Person")]
        [TestCase("UDTPerson", "Person")]
        [TestCase("T_Person", "Person")]
        [TestCase("Table", "Table")]
        [TestCase("Person", "Person")]
        [TestCase("T", "T")]
        [Category("Refactorings")]
        [Category("Extract Udt Class")]
        public void SuggestedClassNameDropsTheTypePrefix(string typeName, string expected)
        {
            Assert.AreEqual(expected, ExtractUdtClassModel.SuggestedClassName(typeName));
        }
    }
}
