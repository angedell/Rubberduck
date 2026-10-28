# Ship: plan

Decisions and reasons are in `docs/adr/0003-ship-builds-a-clean-copy-of-the-saved-workbook.md`; terms in `GLOSSARY.md`.

## What Rubberduck adds to a project (what Ship removes)

| Found in the code | Where | Strip step |
| --- | --- | --- |
| Notes: `@Notes` standard modules holding only comments | VBA project | `RemoveNotes` (whole module) |
| Test modules: `@TestModule`, `@TestMethod`, `@TestInitialize`, `@ModuleInitialize`... and a reference to the Rubberduck unit-testing library | VBA project, references | `RemoveTestModules` |
| `rubberduck.json` | `customUI/rubberduck.json` in the file, with its relationship | `RemoveConfig` |
| `'@Folder`, `'@IgnoreModule`, `'@Description`, `'@EntryPoint`... annotation comments | in every module | `RemoveComments` |
| `Attribute VB_Ext_KEY = "Rubberduck..."` and other descriptive attributes | module headers | `RemoveDescriptiveAttributes` |

Still to check while building, and to be reported in the result if found: staged resource changes pending for the document (`%LOCALAPPDATA%\Rubberduck\PendingResources`, they must not be in the artifact), and any other Rubberduck-written part in the package.

## Pipeline

`ShipPipeline` runs an ordered list of `IShipStep` over a `ShipWorkingSet` (the exported modules as text, the references, the copy's resources). Every step reports a `StepReport` (what it removed, counts). `ShipReport` is the sum plus the problems (modules shipped unchanged because they did not parse), the output path and the chosen file name. A dry run is the preview.

Steps: `RemoveNotes`, `RemoveTestModules` (with its reference), `RemoveComments`, `RemoveDescriptiveAttributes`, `RemoveConfig`. Each can be switched off. No analysis stage exists yet: the ADR names the place.

## Pieces

- **Core** (`Rubberduck.Core/Ship`): `ShipSettings` and its `ship` JSON section, `ShipFolders` (relative paths, `../`, `%ENV%`, known folders, OneDrive, Dropbox), pipeline and steps, `ShipArtifactWriter` (second Excel, late bound, macros disabled, copy deleted after), `ShipService` (rules: saved workbook, Excel only, existing-file choice).
- **Settings file**: `ProjectConfigReader` reads the `ship` section, `ProjectConfigTemplate` writes it with a comment for every setting and lists the folder forms.
- **User setting**: unattended ship, per document, in a small file of the user's own, next to `note-editor.json`.
- **UI**: `ShipDialog` and its view model: two choice cards for the format (`.xlsm` preselected), folder picker with validation, step toggles, summary of what will be removed, confirmation, progress, result with an "open folder" link. Strings in a new `ShipUI.resx`. Command in the Code Explorer's project menu, registered in the IoC installer.
- **API/MCP**: `GET/PUT /projects/{p}/ship/settings`, `GET /projects/{p}/ship/preview`, `POST /projects/{p}/ship`; the same as MCP tools; each added to `Routes()` in `ApiRouter.cs` with the availability condition of the route. The unattended switch is read-only there.

## Tests

Strip steps and the comment stripper are unit-tested without Excel: each edge case (`'`, `Rem`, trailing, continuations with comments, quotes in strings, `#If`, attributes, labels, line numbers), a round trip (the stripped code parses to the same tokens apart from comments), the working project's modules unchanged after a ship, folder resolution, settings read/write. An Excel integration test (category `Excel`, opt-in as in `ExcelIntegrationTests`) ships a real file and opens the result.

## Not exposed or not built

Tree shaking, other hosts, locking the shipped project: deferred, see the issues linked in the final report.

## Overlap with the uncommitted work

These files already have uncommitted changes of others in this branch and Ship needs a small addition to them: `Rubberduck.Main/Root/RubberduckIoCInstaller.cs` (registrations), `Rubberduck.Core/Api/McpEndpoint.cs` and `ApiModels.cs` (tools, models). Strings go in a new resx instead of `CodeExplorerUI.resx`. Ship's own hunks in the shared files will be kept apart from the others when committing.
