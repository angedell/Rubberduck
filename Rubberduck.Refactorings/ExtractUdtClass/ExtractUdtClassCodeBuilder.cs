using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Rubberduck.Refactorings.ExtractUdtClass
{
    /// <summary>
    /// Writes the class for a user defined type. The class keeps its own private copy of the type, so that it also works when the
    /// original is private to another module.
    /// </summary>
    public static class ExtractUdtClassCodeBuilder
    {
        private const string Indent = "    ";
        private const string BackingField = "This";

        public static string ModuleHeader(string className)
        {
            return $@"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""{className}""
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False";
        }

        /// <param name="folderAnnotation">The module's own <c>'@Folder</c> line, to keep the new class next to it, or null.</param>
        public static string ClassBody(string className, string typeName, IReadOnlyList<UdtMemberInfo> members, ExtractUdtClassStyle style, string folderAnnotation = null)
        {
            var fluent = style == ExtractUdtClassStyle.Fluent;
            var code = new StringBuilder();

            if (!string.IsNullOrEmpty(folderAnnotation))
            {
                code.AppendLine(folderAnnotation);
            }
            if (fluent)
            {
                code.AppendLine("'@PredeclaredId");
            }
            code.AppendLine("Option Explicit");
            code.AppendLine();

            code.AppendLine($"Private Type {typeName}");
            foreach (var member in members)
            {
                code.AppendLine($"{Indent}{member.DeclarationText}");
            }
            code.AppendLine("End Type");
            code.AppendLine();
            code.AppendLine($"Private {BackingField} As {typeName}");

            if (fluent)
            {
                code.AppendLine();
                code.AppendLine($"Public Function Create() As {className}");
                code.AppendLine($"{Indent}Set Create = New {className}");
                code.AppendLine("End Function");
            }

            code.AppendLine();
            code.AppendLine($"Public Property Get Self() As {className}");
            code.AppendLine($"{Indent}Set Self = Me");
            code.AppendLine("End Property");

            foreach (var member in members)
            {
                code.AppendLine();
                code.Append(Accessors(member));

                if (fluent && HasWithMethod(member))
                {
                    code.AppendLine();
                    code.Append(FluentMethod(className, member));
                }
            }

            return code.ToString();
        }

        /// <summary>
        /// The members whose names collide with something the class adds itself, or with each other once the class is written: Self, the backing
        /// field This, and for the fluent style Create and the WithX methods (a member WithName next to a member Name). Empty if all is well.
        /// </summary>
        public static IReadOnlyList<string> Clashes(IReadOnlyList<UdtMemberInfo> members, ExtractUdtClassStyle style)
        {
            var fluent = style == ExtractUdtClassStyle.Fluent;
            var problems = new List<string>();

            foreach (var member in members)
            {
                if (member.Name.Equals("Self", StringComparison.OrdinalIgnoreCase))
                {
                    problems.Add($"'{member.Name}' clashes with the Self property of the class.");
                }
                else if (member.Name.Equals(BackingField, StringComparison.OrdinalIgnoreCase))
                {
                    problems.Add($"'{member.Name}' clashes with the private field '{BackingField}' that holds the values.");
                }
                else if (fluent && member.Name.Equals("Create", StringComparison.OrdinalIgnoreCase))
                {
                    problems.Add($"'{member.Name}' clashes with the Create function of the class.");
                }
                else if (fluent)
                {
                    var wither = members.FirstOrDefault(other => HasWithMethod(other)
                        && member.Name.Equals("With" + other.Name, StringComparison.OrdinalIgnoreCase));
                    if (wither != null)
                    {
                        problems.Add($"'{member.Name}' clashes with the With{wither.Name} method that '{wither.Name}' gets.");
                    }
                }
            }

            return problems;
        }

        private static bool HasWithMethod(UdtMemberInfo member) => !member.IsArray && !member.IsUserDefinedType;

        private static string Accessors(UdtMemberInfo member)
        {
            var name = member.Name;

            if (member.IsUserDefinedType)
            {
                // A type that is private to its module cannot cross a procedure boundary of another module, so it is left alone.
                return $"' {name} is a user defined type; wrap it in a class of its own to expose it.{Environment.NewLine}";
            }

            var code = new StringBuilder();
            var value = ParameterName(member);
            var type = member.IsArray ? $"{member.TypeName}()" : member.TypeName;

            code.AppendLine($"Public Property Get {name}() As {type}");
            code.AppendLine(member.IsObject
                ? $"{Indent}Set {name} = {BackingField}.{name}"
                : $"{Indent}{name} = {BackingField}.{name}");
            code.AppendLine("End Property");

            if (member.IsArray)
            {
                return code.ToString();
            }

            code.AppendLine();
            if (member.IsObject)
            {
                code.AppendLine($"Public Property Set {name}(ByVal {value} As {member.TypeName})");
                code.AppendLine($"{Indent}Set {BackingField}.{name} = {value}");
            }
            else
            {
                code.AppendLine($"Public Property Let {name}(ByVal {value} As {member.TypeName})");
                code.AppendLine($"{Indent}{BackingField}.{name} = {value}");
            }
            code.AppendLine("End Property");
            return code.ToString();
        }

        private static string FluentMethod(string className, UdtMemberInfo member)
        {
            var value = ParameterName(member);
            var assignment = member.IsObject
                ? $"Set {BackingField}.{member.Name} = {value}"
                : $"{BackingField}.{member.Name} = {value}";

            var code = new StringBuilder();
            code.AppendLine($"Public Function With{member.Name}(ByVal {value} As {member.TypeName}) As {className}");
            code.AppendLine($"{Indent}{assignment}");
            code.AppendLine($"{Indent}Set With{member.Name} = Me");
            code.AppendLine("End Function");
            return code.ToString();
        }

        private static string ParameterName(UdtMemberInfo member)
        {
            return member.Name.Equals("Value", StringComparison.OrdinalIgnoreCase) ? "NewValue" : "Value";
        }
    }
}
