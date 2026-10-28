using System;
using System.Collections.Generic;

namespace Rubberduck.Api
{
    /// <summary>An error the caller can act on. Becomes an HTTP status and a JSON error body.</summary>
    public class ApiException : Exception
    {
        public ApiException(int status, string code, string message) : base(message)
        {
            Status = status;
            Code = code;
        }

        public int Status { get; }
        public string Code { get; }

        public static ApiException BadRequest(string message) => new ApiException(400, "bad_request", message);
        public static ApiException Unauthorized(string message) => new ApiException(401, "unauthorized", message);
        public static ApiException Forbidden(string code, string message) => new ApiException(403, code, message);
        public static ApiException NotFound(string message) => new ApiException(404, "not_found", message);
        public static ApiException Conflict(string message) => new ApiException(409, "conflict", message);
        public static ApiException PreconditionFailed(string message) => new ApiException(412, "stale_hash", message);
        public static ApiException PreconditionRequired(string message) => new ApiException(428, "hash_required", message);
        public static ApiException Unavailable(string message) => new ApiException(503, "unavailable", message);

        /// <summary>The editor is running or stopped (break mode) code, so the parser cannot work and waiting will not help.</summary>
        public static ApiException NotInDesignMode()
            => new ApiException(409, "vbe_not_in_design_mode",
                "The editor is not in design mode: a macro is running or stopped (break mode). Reset it (Run > Reset, or the Stop button) and retry.");
    }

    /// <summary>Which build of Rubberduck is answering.</summary>
    public class InfoReport
    {
        public string Name { get; set; }

        /// <summary>The full version, e.g. 2.5.92.28. The last number counts up with every build.</summary>
        public string Version { get; set; }

        /// <summary>The start of the git commit the build was made from, if it recorded one.</summary>
        public string Commit { get; set; }

        /// <summary>Version and commit as the splash screen shows them.</summary>
        public string Display { get; set; }

        /// <summary>The API's own version, the one in the URL.</summary>
        public string Api { get; set; }
    }

    public class LogEntry
    {
        /// <summary>As the log writes it: 2026-10-27 06:28:22.9468.</summary>
        public string Time { get; set; }

        /// <summary>Trace, Debug, Info, Warn, Error or Fatal.</summary>
        public string Level { get; set; }

        /// <summary>The class that wrote it, e.g. Rubberduck.Navigation.CodeExplorer.CodeExplorerProjectViewModel.</summary>
        public string Logger { get; set; }

        public string Message { get; set; }
    }

    public class LogReport
    {
        /// <summary>The level in Rubberduck's settings.</summary>
        public string ConfiguredLevel { get; set; }

        /// <summary>The level logging actually runs at. If it differs from the configured one, <see cref="Note"/> says so.</summary>
        public string EffectiveLevel { get; set; }

        /// <summary>The levels there are, most to least verbose.</summary>
        public List<string> Levels { get; set; } = new List<string>();

        /// <summary>The log file, for anybody who wants to open it themselves.</summary>
        public string Path { get; set; }

        public long SizeBytes { get; set; }

        /// <summary>True when the file was longer than what is read, so the oldest entries are not here (look in the archives folder next to the log).</summary>
        public bool StartsMidFile { get; set; }

        public string Note { get; set; }

        /// <summary>How many entries matched the filters, before <c>tail</c> cut them.</summary>
        public int Total { get; set; }

        public bool Truncated { get; set; }

        /// <summary>The latest matching entries, oldest first.</summary>
        public List<LogEntry> Entries { get; set; } = new List<LogEntry>();
    }

    public class ProjectInfo
    {
        public string Id { get; set; }
        public string Name { get; set; }

        /// <summary>The project's file, or null for a document that was never saved. Tells apart projects that share a name.</summary>
        public string Path { get; set; }

        /// <summary>The file's extension (.xlsm, .xlam...), or null.</summary>
        public string Extension { get; set; }
    }

    public class ModuleInfo
    {
        public string Project { get; set; }

        /// <summary>The project's file (null if it was never saved), so the workbook a write went to can be seen in the answer.</summary>
        public string ProjectPath { get; set; }
        public string Name { get; set; }

        /// <summary>standard, class, form or document (a worksheet, workbook or similar object's code).</summary>
        public string Kind { get; set; }

        public int LineCount { get; set; }

        /// <summary>A note (see NoteModule): the module tools leave it alone and the note tools handle it. Set by the store; not part of the answers.</summary>
        [Newtonsoft.Json.JsonIgnore]
        public bool IsNote { get; set; }

        /// <summary>Identifies the code as it is now. Writes have to send it back, so nothing the user typed in the meantime is overwritten.</summary>
        public string Hash { get; set; }

        /// <summary>What was done besides what was asked, e.g. member attributes that the editor dropped and were put back.</summary>
        public List<string> Notes { get; set; }
    }

    public class ModuleContent : ModuleInfo
    {
        /// <summary>The code, without the hidden header (attributes, form layout) the VBE keeps apart.</summary>
        public string Code { get; set; }

        /// <summary>The module as the VBE exports it, header included. Only set when asked for.</summary>
        public string Text { get; set; }

        /// <summary>The hidden attributes as data: <c>{ "module": { "VB_PredeclaredId": true }, "members": { "Greet": { "VB_UserMemId": 0 } } }</c>. Only set when asked for (<c>include=attributes</c>).</summary>
        public Newtonsoft.Json.Linq.JObject Attributes { get; set; }

        /// <summary>Identifies the attributes as they are now (with include=attributes). Hash only covers the code; send this as expectedAttributesHash to change attributes only if nobody else has.</summary>
        public string AttributesHash { get; set; }
    }

    public static class ModuleKinds
    {
        public const string Standard = "standard";
        public const string Class = "class";
        public const string Form = "form";
        public const string Document = "document";

        public static readonly IReadOnlyList<string> Creatable = new[] { Standard, Class, Form };
    }

    public class Diagnostic
    {
        public string Module { get; set; }

        /// <summary>The inspection's name, as Rubberduck shows it.</summary>
        public string Inspection { get; set; }

        /// <summary>The name that goes in an <c>'@IgnoreModule</c> or <c>'@Ignore</c> annotation to silence it.</summary>
        public string AnnotationName { get; set; }

        /// <summary>hint, suggestion, warning or error.</summary>
        public string Severity { get; set; }

        /// <summary>The kind of problem: code quality, maintainability, language opportunity, rubberduck opportunity, etc.</summary>
        public string Type { get; set; }

        public string Description { get; set; }
        public int StartLine { get; set; }
        public int StartColumn { get; set; }
        public int EndLine { get; set; }
        public int EndColumn { get; set; }
    }

    public class DiagnosticsReport
    {
        public string Project { get; set; }

        /// <summary>The module asked for, or null for the whole project.</summary>
        public string Module { get; set; }

        /// <summary>How the parser ended up: Ready, or an error state if some module could not be parsed (the results are then incomplete).</summary>
        public string ParserState { get; set; }

        public int Total { get; set; }

        /// <summary>
        /// "running" when the inspections were not done within the wait: the list is empty and says nothing yet. They go on in the background; ask again
        /// (after <see cref="RetryAfterSeconds"/>) and the answer comes at once. Null when the results are in.
        /// </summary>
        public string Status { get; set; }

        public int? RetryAfterSeconds { get; set; }

        /// <summary>How many results the inspections found of each severity, before any minimum severity or limit was applied.</summary>
        public Dictionary<string, int> CountsBySeverity { get; set; } = new Dictionary<string, int>();

        /// <summary>True when there were more than the limit, and the list was cut.</summary>
        public bool Truncated { get; set; }

        public List<Diagnostic> Diagnostics { get; set; } = new List<Diagnostic>();
    }

    public class ResourceInfo
    {
        public string Name { get; set; }

        /// <summary>ribbon (customUI.xml or customUI14.xml), config (rubberduck.json) or image.</summary>
        public string Kind { get; set; }

        public long Size { get; set; }

        /// <summary>added, modified or removed while the change waits for the document to be closed; otherwise null.</summary>
        public string Pending { get; set; }
    }

    public class ResourceList
    {
        public string Project { get; set; }

        /// <summary>False when the project's document has no ribbon folder: not saved yet, or not an Open XML file.</summary>
        public bool Supported { get; set; }

        public bool PendingChanges { get; set; }
        public string Note { get; set; }
        public List<ResourceInfo> Resources { get; set; } = new List<ResourceInfo>();
    }

    public class ResourceContent : ResourceInfo
    {
        /// <summary>Identifies the content as it is now. Writes and removals have to send it back.</summary>
        public string Hash { get; set; }

        /// <summary>The ribbon XML.</summary>
        public string Xml { get; set; }

        /// <summary>The config's JSON.</summary>
        public string Text { get; set; }

        /// <summary>A picture's bytes.</summary>
        public string ContentBase64 { get; set; }
    }

    public class ResourceChangeResult
    {
        public ResourceInfo Resource { get; set; }
        public string Note { get; set; }

        /// <summary>For rubberduck.json: unknown keys and bad values. The config is stored all the same; Rubberduck ignores what it cannot read.</summary>
        public List<string> Warnings { get; set; }
    }

    public class ImmediateContent
    {
        /// <summary>The Immediate window's text, the last part of it if it is long.</summary>
        public string Text { get; set; }

        public bool Truncated { get; set; }
    }

    public class ImmediateRun
    {
        /// <summary>What appeared in the window after the command(s) were entered, the echo of the commands included.</summary>
        public string Output { get; set; }

        /// <summary>False if the window was still changing when the wait ran out, e.g. the code is still running. Read it again later.</summary>
        public bool Completed { get; set; }

        /// <summary>Where the editor is now: "design", "running_or_break" (the command left a macro running or stopped), or "unresponsive". Same values as in GET /projects.</summary>
        public string VbeMode { get; set; }

        /// <summary>
        /// Errors the editor reported in a dialog while the command ran, e.g. "Run-time error '11': Division by zero".
        /// Rubberduck pressed End or OK for the user, so the editor is not left waiting for a click.
        /// </summary>
        public List<string> Errors { get; set; } = new List<string>();

        /// <summary>A dialog nobody could answer for the agent, such as a message box from the code that ran. A person has to deal with it.</summary>
        public BlockingDialog BlockedByDialog { get; set; }
    }

    public class ParserError
    {
        public string Module { get; set; }
        public int Line { get; set; }
        public int Column { get; set; }
        public string Message { get; set; }

        /// <summary>The text the parser stumbled on.</summary>
        public string Token { get; set; }

        /// <summary>Whether it was the code, or the module's attributes, that could not be parsed.</summary>
        public string CodeKind { get; set; }
    }

    public class ParserErrorsReport
    {
        public string Project { get; set; }
        public string ParserState { get; set; }
        public List<ParserError> Errors { get; set; } = new List<ParserError>();
    }

    public class Symbol
    {
        public string Module { get; set; }
        public string Name { get; set; }

        /// <summary>The kind of declaration as Rubberduck names it: Procedure, Function, PropertyGet, Variable, Constant, UserDefinedType, Enumeration, ...</summary>
        public string Kind { get; set; }

        public string Accessibility { get; set; }

        /// <summary>The declared type, or the return type; empty if none.</summary>
        public string AsType { get; set; }

        /// <summary>The member or module this is declared in.</summary>
        public string Parent { get; set; }

        public int StartLine { get; set; }
        public int EndLine { get; set; }
        public List<string> Annotations { get; set; } = new List<string>();
    }

    public class SymbolsReport
    {
        public string Project { get; set; }
        public string Module { get; set; }
        public int Total { get; set; }
        public bool Truncated { get; set; }
        public List<Symbol> Symbols { get; set; } = new List<Symbol>();
    }

    public class SymbolReference
    {
        public string Module { get; set; }
        public int Line { get; set; }
        public int Column { get; set; }
        public int EndColumn { get; set; }
        public bool IsAssignment { get; set; }

        /// <summary>The member the reference is in.</summary>
        public string In { get; set; }
    }

    public class SymbolUsage
    {
        public Symbol Declaration { get; set; }
        public int Total { get; set; }
        public bool Truncated { get; set; }
        public List<SymbolReference> References { get; set; } = new List<SymbolReference>();
    }

    public class ReferencesReport
    {
        public string Project { get; set; }

        /// <summary>One entry for each declaration with that name; the same name can be declared in several places.</summary>
        public List<SymbolUsage> Symbols { get; set; } = new List<SymbolUsage>();
    }

    public class DefinitionReport
    {
        public bool Found { get; set; }

        /// <summary>The identifier at the position, which may be a use of something declared elsewhere.</summary>
        public string Identifier { get; set; }

        public Symbol Declaration { get; set; }
    }

    public class LibraryReference
    {
        public string Name { get; set; }
        public string Guid { get; set; }
        public string Description { get; set; }
        public int Major { get; set; }
        public int Minor { get; set; }
        public string Path { get; set; }
        public bool IsBuiltIn { get; set; }
        public bool IsBroken { get; set; }
    }

    public class StructureFolder
    {
        /// <summary>The folder from the <c>'@Folder</c> annotations, dot-separated.</summary>
        public string Folder { get; set; }

        public List<string> Modules { get; set; } = new List<string>();
    }

    public class MetricValue
    {
        public string Module { get; set; }

        /// <summary>The member the number is for, or empty for the whole module.</summary>
        public string Member { get; set; }

        public string Metric { get; set; }
        public string Value { get; set; }
    }

    public class TestInfo
    {
        public string Module { get; set; }
        public string Name { get; set; }
        public string Category { get; set; }
        public bool Ignored { get; set; }

        /// <summary>Unknown (not run since the project was loaded), Succeeded, Failed, Inconclusive or Ignored.</summary>
        public string Outcome { get; set; }

        public string Output { get; set; }
        public long DurationMs { get; set; }
    }

    public class TestRunReport
    {
        public string Project { get; set; }

        /// <summary>False if the wait ran out before every test finished; the rest are still running. List the tests later to see them.</summary>
        public bool Completed { get; set; }

        public int Succeeded { get; set; }
        public int Failed { get; set; }
        public int Inconclusive { get; set; }
        public int Ignored { get; set; }
        public List<TestInfo> Tests { get; set; } = new List<TestInfo>();

        /// <summary>
        /// Errors the editor reported in a dialog during the run, e.g. "Compile error: Method or data member not found" for a test that uses something that does not exist.
        /// Rubberduck pressed End or OK for the user, so the run did not sit waiting for a click.
        /// </summary>
        public List<string> Errors { get; set; } = new List<string>();

        /// <summary>A dialog nobody could answer for the agent, such as a message box from the code under test. A person has to deal with it.</summary>
        public BlockingDialog BlockedByDialog { get; set; }
    }

    public class CheckReport
    {
        public string Project { get; set; }
        public string Module { get; set; }

        /// <summary>True when the parser found no error and no inspection found an error or warning (and no test failed, if they ran).</summary>
        public bool Clean { get; set; }

        public ParserErrorsReport ParserErrors { get; set; }
        public DiagnosticsReport Diagnostics { get; set; }

        /// <summary>What the inspections found below the severity that makes the code unclean (hints, suggestions), counted by severity and not listed. Ask for the diagnostics to see them.</summary>
        public Dictionary<string, int> Suppressed { get; set; } = new Dictionary<string, int>();
        public TestRunReport Tests { get; set; }

        /// <summary>"running" when the inspections were not done within the wait, so Clean cannot be known (it is false). Ask again; the results are kept.</summary>
        public string InspectionsStatus { get; set; }
    }

    /// <summary>One change in a batch. <see cref="Op"/> is replaceCode, editLines, replaceText, create or delete.</summary>
    public class ModuleEdit
    {
        public string Op { get; set; }
        public string Module { get; set; }
        public string ExpectedHash { get; set; }
        public string Code { get; set; }
        public string Text { get; set; }
        public string Kind { get; set; }
        public int StartLine { get; set; }
        public int EndLine { get; set; }
        public string Replacement { get; set; }
    }

    public class BatchResult
    {
        public List<ModuleInfo> Modules { get; set; } = new List<ModuleInfo>();
    }
}
