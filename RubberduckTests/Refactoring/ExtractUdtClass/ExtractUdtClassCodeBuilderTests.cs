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
@"Option Explicit

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

            StringAssert.StartsWith("'@PredeclaredId\nOption Explicit\n", actual);
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
            StringAssert.DoesNotContain("With", actual);
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

            StringAssert.StartsWith("'@Folder(\"Model\")\nOption Explicit", actual);
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
