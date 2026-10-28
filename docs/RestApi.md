# API for AI agents

Lets an AI agent read and edit the modules, classes and forms of the projects open in the VBA editor, over HTTP, without exporting anything to a folder. It can also run Rubberduck's code inspections to check what it wrote and, only if you allow it, use the Immediate window. It never touches cells.

## Switching it on

Settings > General > **Let AI agents read and edit modules and forms**. Then **Copy connection details** gives you the addresses and the token to paste into the agent's setup.

- **Also accept agents from the network** is for agents in a sandbox or VM that cannot reach `127.0.0.1` on your machine. The traffic is not encrypted: use it on a network you trust. Windows needs a one-time permission for it, and when it is missing Rubberduck shows a setup window:
  1. It explains why, and shows the exact command with a **Copy** button. **Run as administrator** runs it for you after the Windows permission prompt; or paste it into a Command Prompt opened as administrator.
  2. Only if the sandbox still cannot connect: a second command opens the firewall for private and domain networks. A checkbox adds it to what **Run as administrator** does.
  3. **Try again** starts the API. **Use this PC only** switches the network option off instead.
- **Also let agents read the Immediate window (Debug.Print output), without running anything** and **Also let agents run code** are two separate settings, both off by default; see [the Immediate window](#the-immediate-window).
- **New token** locks out whoever has the current one.
- The first write asks you to allow writes. Yes covers the rest until you switch the API off or ask for a new token.
- Before a module is first changed, and whenever one is replaced or deleted, a copy goes to `%LOCALAPPDATA%\Rubberduck\ApiBackups`.

## MCP

Point an MCP client at `http://127.0.0.1:7437/mcp` with the header `Authorization: Bearer <token>`. Tools: `get_info`, `reparse`, `capture_dump`, `get_log`, `set_log_level`, `list_projects`, `list_modules`, `read_module`, `write_module`, `edit_lines`, `create_module`, `delete_module`, `batch_edit`, `get_diagnostics`, `get_parser_errors`, `check_work`, `list_symbols`, `find_references`, `find_definition`, `list_libraries`, `get_structure`, `get_metrics`, `list_tests`, `list_resources`, `read_resource`, `write_resource`, `delete_resource`, and, only while running code is allowed, `run_tests`, `read_immediate` and `run_immediate`.

The server also offers MCP *resources* (`resources/list`, `resources/read`): `rubberduck://{project}/modules/{module}` (the code), `rubberduck://{project}/diagnostics` and `rubberduck://{project}/parser-errors`, for clients that let the user attach them to a conversation.

## REST

All requests need `Authorization: Bearer <token>`. Bodies and answers are JSON. `{p}` is a project's id or name, `{m}` a module's name.

| Request | Does |
|---|---|
| `GET /api/v1/info` | Which build is answering: `version` (all four parts, e.g. `2.5.92.28`; the last number counts up with every `build.cmd` run), `commit` (the start of the git hash, if the build recorded one), `display` (what the splash screen shows) and `api`. The same build is in the MCP `initialize` answer (`serverInfo.version`, `serverInfo.build`) and the `get_info` tool. Say it when you report a problem. |
| `GET /api/v1/log` | Rubberduck's own log: `configuredLevel` (the setting), `effectiveLevel` (what logging really runs at; a `note` says when they differ or logging is off), `path`, and `entries` (`time`, `level`, `logger`, `message`; a stack trace stays with its entry), oldest first. Query: `tail` (latest entries, default 200, at most 2000), `minLevel` (`Trace`, `Debug`, `Info`, `Warn`, `Error`, `Fatal`), `contains` (in the message or class, any case), `logger` (part of the class name). Only the last 4 MB of the file are read; older entries are in the `archives` folder next to it. |
| `PUT /api/v1/log` | `{ "level": "Debug" }` (`Trace`, `Debug`, `Info`, `Warn`, `Error`, `Fatal` or `Off`). Saved in the user's settings as their own choice (so Rubberduck does not reset a Trace level to Off at shutdown) and applied at once. Needs the same approval as writes. Answers like the read, without entries. |
| `POST /api/v1/reparse` | Does what Refresh in the Code Explorer does: parses everything again, waits until the parser is done, and answers `{ "parserState": "Ready", "projects": [ ... ] }`. A workbook opened after Rubberduck started, or after a restart of the editor, does not show in `GET /projects` until this has run. 409 `vbe_not_in_design_mode` while code is running or stopped. No approval needed: it changes no code. MCP: `reparse`. Always asks for the parse; every other call that needs current results (diagnostics, check, references, tests...) asks for one only when a module's content or a project's references differ from what the parser last finished with, and otherwise answers from what it holds. The log says which: `API catch-up: parse requested because ...` or `... nothing changed ...; no parse requested`. |
| `GET /api/v1/dump` | Whether a memory dump can be taken (`tool`: where Sysinternals ProcDump was found, or null with a `note` on where to put it) and the dumps already in `%LOCALAPPDATA%\Rubberduck\Dumps` (or in the folder the `RUBBERDUCK_DUMP_FOLDER` environment variable names, a shared one say, so the file can be read from another machine). |
| `POST /api/v1/dump` | `{ "kind": "full", "waitSeconds": 120 }`, both optional. Takes a memory dump of the editor's process with ProcDump, for when the editor hangs (calls time out with 503, the parser stays `Busy`, Excel shows "waiting for an OLE action"). It runs on the API's own thread, so it works while the editor's UI thread is stuck. `kind` is `full` (all memory, what reading managed stacks needs) or `mini`. **The dump holds everything in memory, the open documents included**, so it needs the same approval as writes. Answers with `file` and `sizeBytes`. ProcDump is looked for in `RUBBERDUCK_PROCDUMP`, in `%LOCALAPPDATA%\Rubberduck\Tools\procdump.exe` and on the PATH; without it: 404 `procdump_not_found`. MCP: `capture_dump`. |
| `GET /api/v1/projects` | Answers `{ "vbeMode": "design", "projects": [ ... ] }`. `vbeMode` is where the editor is: `design`, `running_or_break` (a macro runs or is stopped: writes, tests and the parser answer 409 `vbe_not_in_design_mode` until it is reset, and the editor cannot say which of the two it is), or `unresponsive` (no answer in 5 seconds, a dialog nobody can see, say; then there is no list, and a `note`). Each project has `id`, `name`, and the file's `path` and `extension` (both absent for a document that was never saved). Names are not unique (two `DevMacros`, say); the `id` addresses one project. It is made from the file's path (`p` and 8 hex digits, e.g. `p3fa9c1d2`: the same file always has the same id, different files never share one), or from the name for a project that was never saved (`u...`). The editor's own id, which older clients stored, still works, but it can be the same for two workbooks (409 then), so use the one listed here. Module answers carry `projectPath`, so the workbook a write went to is in the answer. |
| `GET /api/v1/projects/{p}/modules` | Lists modules: `name`, `kind` (`standard`, `class`, `form`, `document`), `lineCount`. |
| `GET /api/v1/projects/{p}/modules/{m}[?include=text,attributes]` | Reads `code` and `hash` (also the `ETag`). With `include=text`, also the module as exported, header included. With `include=attributes`, the hidden attributes as data: `"attributes": { "module": { "VB_PredeclaredId": true, "VB_Description": "..." }, "members": { "Greet": { "VB_UserMemId": 0 } } }`, plus `attributesHash`: it identifies those attributes the way `hash` identifies the code (`hash` does not change when only an attribute does). |
| `PUT /api/v1/projects/{p}/modules/{m}` | Replaces the code: `{ "code": "...", "expectedHash": "..." }`. Send `"text"` instead of `"code"` to replace the module as exported, header included (a form keeps its layout). |
| `PATCH /api/v1/projects/{p}/modules/{m}` | Replaces lines: `{ "startLine": 3, "endLine": 5, "replacement": "...", "expectedHash": "..." }`. `endLine` = `startLine - 1` inserts. |
| `POST /api/v1/projects/{p}/modules` | Adds a module: `{ "name": "modX", "kind": "standard\|class\|form", "code": "..." }`. Add `"attributes": { ... }` (the shape above) to set hidden attributes as well; the module is created first, so a bad attribute leaves it behind without them. |
| `DELETE /api/v1/projects/{p}/modules/{m}` | Removes it (a document module only loses its code). Needs `If-Match`. |

`expectedHash` may be sent as the `If-Match` header instead.

**Member attributes survive code writes.** The editor drops the hidden attributes of a member (`VB_UserMemId`, `VB_Description`...) when the code holding the member is replaced; the module's own stay. Writes that change code (`code`, `PATCH`, `replaceCode` and `editLines` in a batch) read the attributes first and put back the ones the new code still has a member for, in a second write, and say so in the answer's `notes`. A member that is gone from the new code keeps nothing. If putting them back fails, `notes` says which were lost.

**Hidden attributes without the header.** `PUT` also takes `"attributes"`, alone or together with `"code"`: `{ "attributes": { "module": { "VB_PredeclaredId": true }, "members": { "Greet": { "VB_UserMemId": 0, "VB_Description": "Greets" } } }, "expectedHash": "..." }`. Values are strings, numbers or `true`/`false`; names start with `VB_`; a member has to exist. An attribute that is there is replaced, otherwise it is added. This is what the `Missing...Attribute` and `...OutOfSync` inspections ask for, without reading the export and rebuilding the header yourself. Add `"expectedAttributesHash"` (from a read with `include=attributes`) to an attribute-only `PUT` and it is refused with 412 if the attributes changed since. An attribute-only write also tells the parser to read the module again, so the `Missing...Annotation` and `...OutOfSync` inspections see it at once. It does not read `'@PredeclaredId`-style annotations for you: set the attributes the annotations name.

### Diagnostics

| Request | Does |
|---|---|
| `GET /api/v1/projects/{p}/diagnostics` | Runs Rubberduck's inspections over the project. |
| `GET /api/v1/projects/{p}/modules/{m}/diagnostics` | The same for one module (or `?module=`). |

Every answer has `countsBySeverity`: how many results there were of each severity before `minSeverity` and `limit` were applied. A project-level result (a missing project name, say) has no `module`.

Query: `minSeverity` (`hint`, `suggestion`, `warning`, `error`; leave out anything less severe), `limit` (default 200, maximum 1000) and `waitSeconds` (default 60, 0 to 300).

**Inspections can take minutes.** Rubberduck inspects everything it has parsed, every open project, so on a large one a run takes minutes. It is a background job: the call waits `waitSeconds` for it, and if that is not enough it answers **202** with `"status": "running"`, an empty list and `retryAfterSeconds`; the job carries on, and the next call, once it is done, answers at once. The results are kept for as long as nothing was parsed again (any edit starts a new parse and so a new run). `waitSeconds=0` never waits: it returns the kept results, or starts the job and says "running". The answer lists the worst first, each with `module`, `inspection`, `annotationName` (what goes in `'@Ignore` to silence it), `severity`, `description` and the line and column range. `parserState` is `Ready`, or an error state if a module could not be parsed, in which case the list is incomplete. It asks the parser to catch up with the editor and waits for it, so it can take a few seconds after a write; it needs no approval.

### Batch edits

`POST /api/v1/projects/{p}/edits` with `{ "edits": [ ... ] }`, at most 50, **one edit per module**. Each edit has `op` (`replaceCode`, `editLines`, `replaceText`, `create`, `delete`), `module` and the fields of that operation: `expectedHash` (all but `create`), `code` (`replaceCode`, `create`), `text` (`replaceText`), `startLine`, `endLine`, `replacement` (`editLines`), `kind` (`create`). Everything that can be checked first is: fields, names, and **every hash**. A stale hash in any edit fails the whole batch (412) before the first is applied. Then the edits are applied in order under one approval. If one still fails, the error names the modules already changed (their old versions are in the backup folder) and says nothing after it was applied. The batch is not atomic beyond that. Answer: `{ "modules": [ ... ] }`.

### Understanding the code

All read-only and without approval. Each waits for the parser to catch up with the editor.

| Request | Does |
|---|---|
| `GET /api/v1/projects/{p}/parser-errors` | The code the parser could not understand: `module`, `line`, `column`, `message`, `token`, `codeKind`. Inspections of such code are incomplete. |
| `GET /api/v1/projects/{p}/symbols?module=M` | What the module declares, with `kind`, `asType`, `accessibility`, `startLine`, `endLine`, `annotations`. `includeLocals=true` adds parameters and local variables; `kind=` filters. |
| `GET /api/v1/projects/{p}/symbols/references?name=N` | Every use of a declared name, per declaration (`module=` narrows, `limit=` default 200, maximum 2000). |
| `GET /api/v1/projects/{p}/symbols/definition?module=M&line=L&column=C` | Where the identifier at a position (1-based) is declared. `found` is `false` for things outside the project's own code. |
| `GET /api/v1/projects/{p}/libraries` | The referenced libraries: name, version, path, `isBuiltIn`, `isBroken`. |
| `GET /api/v1/projects/{p}/structure` | The modules by `'@Folder`. |
| `GET /api/v1/projects/{p}/metrics` | Code metrics per module and member (`module=` narrows). |

### Unit tests

| Request | Does |
|---|---|
| `GET /api/v1/projects/{p}/tests` | Lists the `'@TestMethod` procedures (`module=` narrows) with the outcome of their last run since Rubberduck started (`Unknown` if none). Runs nothing, needs no opt-in. |
| `POST /api/v1/projects/{p}/tests/run` | `{ "module": "...", "test": "...", "waitSeconds": 60 }`, all optional (`test` needs `module`). Runs them and answers with counts and each test's `outcome` and `output`. `completed: false` if the wait (at most 300 s) ran out first. One run at a time (409 otherwise). |

**Running tests executes your VBA.** It needs **Also let agents run code** switched on (otherwise 403 `execute_disabled`) and shares the Immediate window's approval: the first run, or the first command, asks you once. The editor has to be in design mode with the parser ready, otherwise 503.

### Check my work

`POST /api/v1/projects/{p}/check` with `{ "module": "...", "runTests": false }` (both optional). Parser errors first; if the code parsed, the inspections of `warning` or `error` severity; and with `runTests`, the unit tests (same rules as above). `clean` is `true` when there is nothing to fix. It takes `waitSeconds` too (as above); when the inspections are not done in time, `inspectionsStatus` is `"running"` and `clean` is `false`: ask again, and the results are there. It looks at warnings and errors only; `suppressed` counts what was left out (`{ "suggestion": 4, "hint": 1 }`), so a clean answer is not mistaken for "the inspections found nothing". Ask for the diagnostics to see those. Code that did not parse gets no inspections or tests: fix the parser errors first.

### Ribbon XML and pictures

The ribbon XML and pictures stored inside the project's document file: `customUI.xml` (Office 2007) or `customUI14.xml` (2010 and later), and the pictures in `customUI/images` (png, jpg, gif, bmp, ico; 5 MB at most). Only saved Open XML documents have them; for others `supported` is `false`. `{r}` is the resource's file name.

| Request | Does |
|---|---|
| `GET /api/v1/projects/{p}/resources` | Lists them: `name`, `kind` (`ribbon`, `image`), `size`, `pending`. |
| `DELETE /api/v1/projects/{p}/resources` | Discards every change that is staged and not yet in the file (what the Code Explorer's discard does). Answers with the list. Needs the same approval as writes, unless nothing is staged. |
| `GET /api/v1/projects/{p}/resources/{r}` | Reads one with its `hash` (also the `ETag`): `xml` for ribbon XML, `contentBase64` for a picture. |
| `PUT /api/v1/projects/{p}/resources/{r}` | Adds or replaces: `{ "xml": "...", "expectedHash": "..." }` or `{ "contentBase64": "...", "expectedHash": "..." }`. The hash is only needed when the resource exists. |
| `DELETE /api/v1/projects/{p}/resources/{r}` | Removes a picture (ribbon XML can only be replaced). Needs `If-Match`. |

The document is open, so Office has its file locked: a write is **staged** and goes into the file when the document is closed (`pending` is `added`, `modified` or `removed` until then). Reads already include staged changes. The root of the XML must be `<customUI>` in the namespace that goes with the file name, or the write is refused with 400 and the line of the problem. Pictures are referred to in the XML by file name without extension, e.g. `image="logo"` for `logo.png`. Writes need the same approval as module writes.

#### rubberduck.json

The project's `rubberduck.json` lives in the same list, as the resource named `rubberduck.json` (`kind` `config`), stored in the document at `customUI/rubberduck.json`. If the document has none, Rubberduck creates one (staged) the first time the project is shown. It is text: read it from `text`, write it with `{ "text": "...", "expectedHash": "..." }`. Malformed JSON is refused with 400 and the position; settings that are wrong but parse (an unknown value, say) are accepted and returned in `warnings`. It cannot be deleted. The same applies to the MCP tools `list_resources`, `read_resource`, `write_resource` (with `text`) and `delete_resource`.

### The Immediate window

Two settings (General), both off by default:

| Setting | Allows | Otherwise |
|---|---|---|
| **Also let agents read the Immediate window** | `GET /immediate` and `GET /immediate/output` (MCP: `read_immediate`, `read_debug_output`). Reads only, no approval dialog. What the window shows can include values from your data. | 403 `immediate_disabled` |
| **Also let agents run code** | `POST /immediate` (MCP: `run_immediate`) and running tests, after your approval. **Includes reading**, so it needs no second setting. | 403 `immediate_disabled`; with only reading allowed, the message says so |

With only reading allowed, an agent can still add `Debug.Print` lines to the code, but has to ask you to run it. The MCP instructions say so, and they leave the Debug.Print advice out when reading is off.

| Request | Does |
|---|---|
| `GET /api/v1/immediate` | `{ "text": "..." }`: what the window shows (the last 20,000 characters). |
| `POST /api/v1/immediate` | `{ "command": "?1+1", "waitMs": 5000 }`: enters the line(s) and runs them. Answers `{ "output": "...", "completed": true, "vbeMode": "design" }`. `vbeMode` is where the editor is after the command (as in `GET /projects`): `running_or_break` means the command left a macro running or stopped (`Stop`, `Debug.Assert False`), and writes, tests and the parser will answer 409 until the user resets it. |
| `GET /api/v1/immediate/output?since=&limit=&contains=&waitMs=` | What was printed to the window, `Debug.Print` included, **also while nobody was asking** (a form event, a timer, a macro the user started). Answers `{ "text": "...", "lineCount": 3, "next": 42, "more": false, "missed": false }`. Pass `next` as `since` on the next call to get only what is new; without `since` you get the newest `limit` lines (default 200, at most 1,000). `contains` keeps lines with that text (any case), `waitMs` (at most 30,000) waits for output when there is none yet. MCP tool: `read_debug_output`. Runs nothing, so it needs no approval. |


**How the output is captured.** The editor has no event for `Debug.Print`, so while the API is on and the Immediate window is allowed, Rubberduck looks at the window twice a second and keeps what was added as numbered lines (the newest 5,000). Text that was in the window before Rubberduck's first look is not included (`GET /immediate` still shows it). `missed` is true when lines after your `since` were lost: the buffer wrapped, or the window scrolled out more than it holds between two looks. A closed Immediate window is not opened by the background look (only a request does that), so output printed while it is closed is picked up when it is next open. The lines you or an agent typed into the window are part of its text and show up too. A line the user is still typing is not taken for output until it has stayed the same for a look.

**Debugging with `Debug.Print`.** The MCP instructions encourage agents to make code report on itself instead of guessing: put `Debug.Print "[Module.Procedure] x=" & x` lines in the code, wrapped in a conditional compilation constant, run it, and read the lines back:

```vb
#Const DebugLog = True

#If DebugLog Then
    Debug.Print "[Orders.Total] qty=" & qty & " price=" & price
#End If
```

`#Const` is private to its module, so it is switched on per module; the agent is asked to set it to `False` (or remove the lines) when it is done and to tell the user.

**Entering a command runs it**, with your rights: `?expression` prints a value, a statement runs, a procedure name calls it. The first command shows a dialog (separate from the one for writes) that you have to accept. `completed: false` means the window was still changing when `waitMs` (at most 60,000) ran out, for example because the code is still running: read the window again later. Commands are limited to 50 lines and 10,000 characters, and to characters the editor's window can take (no Unicode beyond the ANSI code page). The editor has to be open; the window is shown if it was hidden.

**Error dialogs.** Code that fails shows the editor's error dialog, which nobody but a human could click. While a command runs, these dialogs are looked for: the editor's own error dialog (it has a Help button and belongs to the editor window) is closed with **End**, never **Debug**, so the editor does not stay in break mode, and its text is added to the answer's `errors`. Any other dialog (a message box of your code, Excel's own) is left alone: the answer then has `completed: false` and `blockedByDialog` with its `title`, `message` and `buttons`, and the user has to answer it. Finding and closing the error dialog was tried on an English Office 365 editor; the dialog's buttons are matched by their text.

**What is logged.** Every request gets a line in Rubberduck's log (`GET /api/v1/log?logger=ApiServer`): method, path, status and duration, at Debug level, and at Info when it took 2 seconds or more. An edit that needs the parser paused logs when it asks and how the pause ended (`logger=VbeModuleStore`).

Errors look like `{ "error": { "code": "...", "message": "..." } }`:

| Status | Meaning |
|---|---|
| 401 / 403 | Wrong or missing token; a browser request; network access off; writes not allowed by the user. |
| 412 | The module changed since you read it. Read it again. |
| 422 | The code contains `Attribute VB_...` lines (send it without its header, or use `text`); or a command has a character the Immediate window cannot take. |
| 428 | No hash sent. |
| 409 | `vbe_not_in_design_mode`: a macro is running or stopped (break mode), so the parser cannot work. Reset it (Run > Reset) and retry; waiting does not help. Also: a project name is ambiguous (use the id). |
| 503 | Rubberduck is still parsing or the editor did not answer in time. Retry after the seconds in the `Retry-After` header. Edits are possible while the parser is in an error state, so a module that does not parse can be fixed or deleted. |

## Example

```
curl -H "Authorization: Bearer $TOKEN" http://127.0.0.1:7437/api/v1/projects/VBAProject/modules/Module1
curl -X PUT -H "Authorization: Bearer $TOKEN" -d '{"code":"Option Explicit\r\nSub Hello()\r\nEnd Sub","expectedHash":"<hash from the read>"}' \
     http://127.0.0.1:7437/api/v1/projects/VBAProject/modules/Module1
```

The reasoning behind the design is in `docs/adr/0001-local-api-for-ai-agents.md`.
