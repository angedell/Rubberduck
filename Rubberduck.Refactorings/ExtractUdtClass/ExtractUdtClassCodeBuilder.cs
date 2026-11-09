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

        /// <summary>The header the VBE reads when it imports the module, with the line endings it insists on: a bare line feed makes it show the header as code.</summary>
        public static string ModuleHeader(string className)
        {
            return string.Join("\r\n", new[]
            {
                "VERSION 1.0 CLASS",
                "BEGIN",
                "  MultiUse = -1  'True",
                "END",
                $"Attribute VB_Name = \"{className}\"",
                "Attribute VB_GlobalNameSpace = False",
                "Attribute VB_Creatable = False",
                "Attribute VB_PredeclaredId = True",
                "Attribute VB_Exposed = False"
            });
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
            code.Append(UsageComment(className, members, style));
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
                code.Append(CreateFunction(className, members));
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
        /// Comment lines for the top of the class that show how to use it: how to make one and set and read its members. The values are samples of the
        /// member's type. Arrays and nested types, which have no setter, are left out.
        /// </summary>
        public static string UsageComment(string className, IReadOnlyList<UdtMemberInfo> members, ExtractUdtClassStyle style)
        {
            const string variable = "item";
            var settable = members.Where(member => !member.IsArray && !member.IsUserDefinedType).ToList();
            var lines = new List<string> { "Usage:", $"  Dim {variable} As {className}" };

            if (style == ExtractUdtClassStyle.Fluent)
            {
                var required = RequiredMembers(members);
                var chained = settable.Where(member => !required.Contains(member)).ToList();
                var create = $"{className}.Create({string.Join(", ", required.Select(SampleValue))})";

                lines.Add($"  Set {variable} = {create}{(chained.Count > 0 ? " _" : string.Empty)}");
                for (var i = 0; i < chained.Count; i++)
                {
                    lines.Add($"      .{WithMethodName(chained[i])}({SampleValue(chained[i])}){(i < chained.Count - 1 ? " _" : string.Empty)}");
                }
            }
            else
            {
                lines.Add($"  Set {variable} = New {className}");
                foreach (var member in settable)
                {
                    lines.Add(member.IsObject
                        ? $"  Set {variable}.{member.Name} = {SampleValue(member)}"
                        : $"  {variable}.{member.Name} = {SampleValue(member)}");
                }
            }

            var read = settable.FirstOrDefault(member => !member.IsObject);
            if (read != null)
            {
                lines.Add($"  Debug.Print {variable}.{read.Name}");
            }
            lines.Add("Self returns the object itself, which suits With blocks and chaining.");

            var code = new StringBuilder();
            foreach (var line in lines)
            {
                code.AppendLine($"'{(line.StartsWith(" ") ? string.Empty : " ")}{line}");
            }
            return code.ToString();
        }

        private static string SampleValue(UdtMemberInfo member)
        {
            if (member.IsObject)
            {
                return "Nothing";
            }

            switch (member.TypeName.ToUpperInvariant())
            {
                case "STRING":
                case "VARIANT":
                    return "\"text\"";
                case "BOOLEAN":
                    return "True";
                case "DATE":
                    return "#1/1/2000#";
                default:
                    //The numeric types, and enumerations.
                    return "0";
            }
        }

        /// <summary>The members that Create takes as arguments, in the order of the type: the ones marked required that can be.</summary>
        public static IReadOnlyList<UdtMemberInfo> RequiredMembers(IReadOnlyList<UdtMemberInfo> members)
            => members.Where(member => member.IsRequired && member.CanBeRequired).ToList();

        private static string CreateFunction(string className, IReadOnlyList<UdtMemberInfo> members)
        {
            var required = RequiredMembers(members);
            var code = new StringBuilder();

            if (required.Count == 0)
            {
                code.AppendLine($"Public Function Create() As {className}");
                code.AppendLine($"{Indent}Set Create = New {className}");
                code.AppendLine("End Function");
                return code.ToString();
            }

            //The new object cannot be called Create's own name, nor the name of one of the arguments.
            var created = "NewInstance";
            while (required.Any(member => member.Name.Equals(created, StringComparison.OrdinalIgnoreCase)))
            {
                created += "_";
            }

            var arguments = string.Join(", ", required.Select(member => $"ByVal {member.Name} As {member.TypeName}"));
            code.AppendLine($"Public Function Create({arguments}) As {className}");
            code.AppendLine($"{Indent}Dim {created} As {className}");
            code.AppendLine($"{Indent}Set {created} = New {className}");
            foreach (var member in required)
            {
                code.AppendLine(member.IsObject
                    ? $"{Indent}Set {created}.{member.Name} = {member.Name}"
                    : $"{Indent}{created}.{member.Name} = {member.Name}");
            }
            code.AppendLine($"{Indent}Set Create = {created}");
            code.AppendLine("End Function");
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
                        && member.Name.Equals(WithMethodName(other), StringComparison.OrdinalIgnoreCase));
                    if (wither != null)
                    {
                        problems.Add($"'{member.Name}' clashes with the {WithMethodName(wither)} method that '{wither.Name}' gets.");
                    }
                }
            }

            return problems;
        }

        /// <summary>WithName for a member called name: the member's first letter is capitalised, as the VBE would not do it for us.</summary>
        private static string WithMethodName(UdtMemberInfo member) => "With" + char.ToUpperInvariant(member.Name[0]) + member.Name.Substring(1);

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
            var method = WithMethodName(member);
            code.AppendLine($"Public Function {method}(ByVal {value} As {member.TypeName}) As {className}");
            code.AppendLine($"{Indent}{assignment}");
            code.AppendLine($"{Indent}Set {method} = Me");
            code.AppendLine("End Function");
            return code.ToString();
        }

        private static string ParameterName(UdtMemberInfo member)
        {
            return member.Name.Equals("Value", StringComparison.OrdinalIgnoreCase) ? "NewValue" : "Value";
        }
    }
}
