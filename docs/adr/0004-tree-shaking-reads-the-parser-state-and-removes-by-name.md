# 0004. Tree shaking reads the parser state and removes by name

## Status

Proposed.

## Context

Ship (ADR 0003) strips what Rubberduck adds, but ships all of the user's code. Tree shaking removes code that cannot be reached. Ship's strip steps work on the exported text of each module, parsed in isolation: that knows nothing across modules. Reachability needs resolved references, which only the parser state of the open project has, and that state parses the code pane, whose line numbers differ from the exported file (header and attributes). The annotations that name roots (`@EntryPoint`, `@ExcelHotkey`) are comments, which the comment step removes.

## Decision

1. **An analysis stage, then a step.** `IShipAnalysis` runs first, on the parser state, while annotations are still in the text, and returns *what to remove by name* (module, member, kind), the roots with the reason each is one, and the dynamic calls it cannot follow. The `treeShaking` step then removes those names from the exported text with the same parser-based edits as the comment step (the parse tree's own spans), so the two never disagree about line numbers. The analysis changes nothing.
2. **Roots.** Conservative mode (default): every public procedure of a standard module, `@EntryPoint`/`@ExcelHotkey` members, event handlers (document, chart, UserForm and control events, `WithEvents`), interface implementations, `Class_Initialize`/`Class_Terminate`, `Auto_Open`/`Auto_Close`, public members of classes with `VB_Exposed = True`. `entryPointsOnly` mode: no blanket rule for public procedures; only annotations, the `keep` list, the always-kept code above and the detected roots below.
3. **Detected roots (both modes).** A member name as a whole word anywhere inside a string literal (`"=MyUdf(A1)"`, `Evaluate("MyUdf(2)")`, `"'Module1.Foo 1'"`), including the literal that is the name, and the targets of `Application.Run`, `OnTime`, `OnKey`, `CallByName`; ribbon callback names in `customUI.xml`; macro names on shapes and buttons, defined names and UDF names in formulas, found by reading the document's parts (ribbon, sheets, drawings, tables, pivots, query tables) for the name as a whole word; every identifier inside an `#If`/`#Else` branch of any module is a name root too, since the parser sees only the active branch. Deliberately over-cautious: a name that only appears in a string keeps the member. Calls whose target is computed are listed in the preview and the result as dynamic calls.
4. **What goes.** Unreachable procedures and functions, module-level variables, constants, `Type`s, `Enum`s and `Declare`s, and a standard or class module left empty that nothing references by name. Never removed: UserForms; document modules (only emptied of unreachable members); `Option` and `Attribute` lines.
5. **All or nothing.** If the parser is not ready or any module does not parse, tree shaking is skipped for the whole project and the reason is reported; the other steps still run. The reference graph spans modules, so a module-by-module result would be unsafe.
6. **Off by default**, because it is the one step that can remove code the user needs. The preview lists every removal by name. `ship.treeShaking.mode` and `ship.treeShaking.keep` are in `rubberduck.json`; agents can edit them like the other settings.
7. **No compile as proof.** The check is a corpus test (every kept member's references are kept, directly or transitively) and a re-parse of the shaken text. Compiling the copy in the hidden Excel would catch more but a compile error raises a VBE modal that Excel does not suppress; it is left for a later decision.

Rejected: a regex or name-only dead-code pass on the exported text (no resolution: two modules with a `Run` would be confused); removing by the parser state's own line numbers (they do not match the exported file); a project-wide `Application.Run` wrapper class that Rubberduck knows, with an inspection forbidding other `Application.Run` calls (it cannot see a computed name either, `Application.Run`, `OnTime` and `OnKey` can only reach public standard-module procedures, which conservative mode keeps anyway, and it would change every project that adopts it).

## Consequences

- Ship now needs the project to be parsed: it asks the parser to catch up and waits, like the API does, and reports "not ready" instead of guessing.
- The preview and the result list non-literal dynamic calls, and agents are told in the MCP instructions to use literal names and to annotate or `keep` what they call dynamically.
- Strings that happen to equal a member's name keep it; shipping a smaller file needs `entryPointsOnly` and a `keep` list.
- The seam named in ADR 0003 is used as described there: one analysis stage before the steps, one step after it.
