using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    /// <summary>
    /// @Notes annotation, marks a standard module as a note: markdown text kept in the project, with a comment mark in front of every line so that it compiles.
    /// Rubberduck shows it in the Code Explorer under its name, opens it in a markdown editor, leaves it out of inspections, and offers it to AI agents through its own tools instead of the module tools.
    /// </summary>
    /// <parameter name="Name" type="Text">
    /// Optional. The name the note is shown under, e.g. "CLAUDE.md"; VBA module names cannot contain a dot. Without it the note is shown as the module name followed by ".md".
    /// </parameter>
    /// <remarks>
    /// The module must hold nothing but comments and blank lines. A module with code is an ordinary module whatever its annotations say.
    /// Combine with @Folder to make the note guidance for the modules in that folder and its subfolders.
    /// </remarks>
    /// <example>
    /// <module name="CLAUDE_md" type="Standard Module">
    /// <![CDATA[
    /// '@Notes("CLAUDE.md")
    /// '@Folder("Api.Core")
    /// ' # Rules for this folder
    /// '
    /// ' - Every public member gets a test.
    /// ]]>
    /// </module>
    /// </example>
    public sealed class NotesAnnotation : AnnotationBase
    {
        public NotesAnnotation()
            : base("Notes", AnnotationTarget.Module, 0, 1, new[] { AnnotationArgumentType.Text })
        {}

        public override ComponentType? RequiredComponentType => ComponentType.StandardModule;
    }
}
