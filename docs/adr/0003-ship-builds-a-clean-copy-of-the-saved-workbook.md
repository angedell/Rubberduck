# 0003. Ship builds a clean copy from the saved workbook, in a second Excel

## Status

Proposed.

## Context

Ship (see the glossary) writes an `.xlam` or `.xlsm` without what Rubberduck added for development. Rubberduck sits inside the VBE of the user's Excel and can export, import and edit modules, but it cannot write another file's VBA project: that is a compressed binary inside the document. The open project must never be changed by Ship.

## Decision

1. **Work on a copy of the file.** The saved workbook file is copied to a temp folder. A hidden second Excel instance (late bound, `Excel.Application`) opens the copy with macros force-disabled, so no user code runs, and the strip steps are applied to the copy's project. The copy is saved as the artifact (`.xlsm`, or `.xlam` through `SaveAs`) and deleted. A workbook with unsaved changes, or one that was never saved, is rejected: the copy is of the file, not of the live state.
   Only the copy's in-memory project is edited, never the user's. Stripped text is produced outside the VBE (from an export of the user's project, see 2) and replaces a module's code in the copy.
2. **Strip text with the parser, not a regex.** Each module is exported to the temp folder and parsed with the existing `IStringParser`. Comments are removed by deleting the parser's comment contexts through a `TokenStreamRewriter`, so string literals, `Rem`, trailing comments, line continuations, `#If` lines, labels and line numbers keep their meaning. A module that does not parse ships unchanged and is listed in the result. Parse errors are shown, not blocking.
3. **Strip steps are a pipeline.** Each step is a small class taking the working set (modules and the copy's resources) and reporting what it removed: notes (whole modules), test modules and the Rubberduck unit-testing reference, comments (this includes `'@` annotations), descriptive attributes, `rubberduck.json`. The result report is the sum of the steps' reports and is also the preview, produced by running the pipeline without writing.
4. **Attributes.** Functional attributes ship: `VB_PredeclaredId`, `VB_Exposed`, `VB_Creatable`, `VB_GlobalNameSpace`, `VB_UserMemId`, `VB_VarUserMemId`, `VB_MemberFlags`, `VB_Name`, and `VB_Description` and `VB_ProcData.VB_Invoke_Func`/`_PropertyPut`/`_PropertyPutRef` on public procedures of standard modules that are not under `Option Private Module`. All others, including every `VB_Ext_KEY`, are dropped. The shipped project is not meant to be read, so descriptive attributes are not kept for people.
5. **Ribbon XML and pictures ship**; only `customUI/rubberduck.json` is removed.
6. **Settings** are the `ship` section of the project's `rubberduck.json` (shared, staged like any resource change, see ADR 0002). The switch for **unattended ship** is not there: an agent can edit the project's JSON, so a switch that waives the agent's approval must live in the user's own settings, set from the Ship dialog only.
7. **Agents.** Ship is on the REST API and MCP (preview, settings, run). Run needs write approval and `confirm: true`, unless the user switched unattended ship on. Because an agent can edit the folder and file name in the settings, three limits hold for everyone: the shipped file is never the workbook itself, never goes into a folder that Windows or Excel loads files from by itself (Startup, XLSTART), and an output folder can use only environment variables that name a folder (so a secret cannot be read back through the resolved path). An agent cannot choose `replace` for an existing file unless the user switched unattended ship on; replacing is otherwise done from the dialog. A replaced file stays as `.bak` (`.bak2`... if that exists), and Ship never overwrites a backup.
8. **Threads.** The dialog and the API run the service off the editor's thread (as the parser already does for the VBE), so the watchdog on the second Excel can fire. The API only finds the project on the editor's thread.

Rejected: (a) a fresh workbook with imported modules, which loses sheets, data, forms' binaries and the ribbon; (b) stripping the user's live project and undoing it, which breaks the rule that Ship never touches the working project; (c) rewriting `vbaProject.bin` ourselves, fragile and unsupported.

## Seam for tree shaking (built in ADR 0004)

The pipeline is an ordered list of steps over a working set. Tree shaking will be added as one more step placed after an analysis step that reads parser state (entry points and the declaration graph) and marks declarations. In v1 nothing for it existed; ADR 0004 adds the analysis stage and the step exactly there.

## Consequences

- Ship needs Excel with "Trust access to the VBA project object model" (as Rubberduck does) and works on Excel projects only; other hosts have no Ship command.
- Unsaved work must be saved first. Locking the shipped project with a password is out of scope: the object model has no call for it.
- `ProjectConfigReader` must learn the `ship` section; today it reports unknown sections.
- Dropping descriptive attributes loses Object Browser and Insert Function text outside the rule in 4.
