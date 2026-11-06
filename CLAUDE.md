## Agent skills

### Issue tracker

Issues live in GitHub Issues (angedell/Rubberduck), via the `gh` CLI. See `docs/agents/issue-tracker.md`.

### Triage labels

Default vocabulary: `needs-triage`, `needs-info`, `ready-for-agent`, `ready-for-human`, `wontfix`. See `docs/agents/triage-labels.md`.

### Domain docs

Single-context: one `GLOSSARY.md` + `docs/adr/` at the repo root. See `docs/agents/domain.md`.

## REST API

### Route index

`GET /api/v1` lists the routes, from `ApiRouter.Routes()` in `Rubberduck.Core/Api/ApiRouter.cs`. The list is written by hand, so when you add, rename or remove a route in `ApiRouter.Route`, `RouteProject` or one of the routers it hands off to (quickfix, refactor), update `Routes()` in the same change. A forgotten route is simply missing from the index and no test notices. Keep its availability condition the same as the one on the route itself (for example `_notes != null`).

### Split Parameters

The Split Parameters refactoring (`Rubberduck.Refactorings/SplitParameters`) puts every parameter of a signature on its own line, with leading commas and the closing parenthesis on its own line, so that signature changes are one-line diffs. It works on one member (`Refactor > Split Parameters`) or on every member of a project in one rewrite (`Split Parameters in Project`). Over the API it is `POST /api/v1/projects/{p}/refactor/split-parameters[/preview]` and the MCP tools `preview_split_parameters` and `split_parameters` (`Rubberduck.Core/Api/RefactorSplitParametersApi.cs`); its route is in `ApiRouter.Routes()`, and the tools are listed in `docs/RestApi.md`. The indenter keeps that layout by fixed offsets (`LogicalCodeLine.IsLeadingCommaSignature`), so reindenting does not realign it to the opening parenthesis.
