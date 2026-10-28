using System.Collections.Generic;

namespace Rubberduck.Api
{
    /// <summary>
    /// The modules and forms of the open projects, as the API sees them. Implementations throw <see cref="ApiException"/> for what the caller got wrong.
    /// Every write takes the hash from the read it is based on and fails with 412 if the code changed since.
    /// </summary>
    public interface IModuleStore
    {
        IReadOnlyList<ProjectInfo> Projects();

        IReadOnlyList<ModuleInfo> Modules(string project);

        ModuleContent Read(string project, string module, bool includeText);

        /// <summary>Replaces the code of an existing module. The hidden header, and a form's layout, stay as they are.</summary>
        ModuleInfo ReplaceCode(string project, string module, string code, string expectedHash);

        /// <summary>Replaces lines <paramref name="startLine"/> to <paramref name="endLine"/> (1-based, inclusive). An end line one before the start inserts.</summary>
        ModuleInfo EditLines(string project, string module, int startLine, int endLine, string replacement, string expectedHash);

        /// <summary>Replaces the whole module, header included, as an import of the VBE's own export format. A form keeps its binary layout.</summary>
        ModuleInfo ReplaceText(string project, string module, string text, string expectedHash);

        ModuleInfo Create(string project, string module, string kind, string code);

        void Delete(string project, string module, string expectedHash);
    }
}
