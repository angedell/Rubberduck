using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.ExtractUdtClass
{
    /// <summary>What the generated class needs to know about one member of the user defined type.</summary>
    public class UdtMemberInfo
    {
        public UdtMemberInfo(string name, string typeName, bool isArray, bool isObject, bool isUserDefinedType, string declarationText)
        {
            Name = name;
            TypeName = typeName;
            IsArray = isArray;
            IsObject = isObject;
            IsUserDefinedType = isUserDefinedType;
            DeclarationText = declarationText;
        }

        public string Name { get; }
        /// <summary>The type without any array designator.</summary>
        public string TypeName { get; }
        public bool IsArray { get; }
        public bool IsObject { get; }
        public bool IsUserDefinedType { get; }
        /// <summary>The member as it is written in the type, for the copy of the type in the new class.</summary>
        public string DeclarationText { get; }

        /// <summary>An array cannot be handed over as a single value, and a Type inside the Type is not wrapped at all, so neither can be an argument of Create.</summary>
        public bool CanBeRequired => !IsArray && !IsUserDefinedType;

        /// <summary>The member has to be given when the object is made: it becomes an argument of Create. Only the fluent style has a Create.</summary>
        public bool IsRequired { get; set; }

        public static UdtMemberInfo From(Declaration member)
        {
            var context = member.Context;
            var text = context?.Start?.InputStream != null && context.Stop != null
                ? context.Start.InputStream.GetText(Interval.Of(context.Start.StartIndex, context.Stop.StopIndex))
                : $"{member.IdentifierName} As {member.AsTypeName}";

            return new UdtMemberInfo(
                member.IdentifierName,
                member.AsTypeNameWithoutArrayDesignator,
                member.IsArray,
                member.IsObject,
                member.AsTypeDeclaration?.DeclarationType == DeclarationType.UserDefinedType,
                text);
        }
    }
}
