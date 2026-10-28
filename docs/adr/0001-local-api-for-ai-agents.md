# 0001. A local HTTP API and MCP server for AI agents

## Status

Accepted.

## Context

AI agents that write VBA have had to work through a folder: export the modules, edit the files, import them back (Git sync). That needs a folder both sides can reach, which sandboxed agents often do not have. They can usually reach a network address.

An API that can write code into a workbook is a risk of its own. VBA is executable, and the server lives inside the user's Office process.

## Decision

Rubberduck runs an HTTP server in the add-in, **off by default**, switched on in Settings > General. It offers the same operations twice: REST under `/api/v1` and an MCP server (Streamable HTTP, JSON responses) at `/mcp`. Both go through one class, `ModuleApi`, so they enforce the same rules.

V1 covers modules, classes, forms and document modules (a sheet's or the workbook's *code*). It never touches cells. It also offers Rubberduck's code inspections (diagnostics, read-only). By default it never runs code; the Immediate window is a separate, stricter opt-in (see below).

Security model:

1. **Bearer token.** Every request needs it. It is generated once and kept in `%APPDATA%\Rubberduck\api-token.txt`, so it can be given to a sandbox once. The user can regenerate it in Settings, which also takes back the write approval.
2. **No browsers.** Any request with an `Origin` header is refused. Agents do not send one; web pages always do. This stops cross-site requests and DNS rebinding from a page on the user's machine.
3. **Loopback by default.** The server listens on `127.0.0.1` and `localhost` and refuses other `Host` values. A separate toggle makes it listen on every address (`http://+:port/`), for agents in a sandbox or VM. Windows needs a one-time `netsh http add urlacl` for that; when it is missing, a setup window explains why, shows the commands with Copy buttons, and can run them elevated. The traffic is plain HTTP, so this is for networks the user trusts. The token is the only protection on it.
4. **Writes need approval, once.** Reads work with the token. The first write after the API starts shows a dialog; yes allows writes until the API is switched off or the token regenerated. A no is remembered for a minute so an agent cannot flood the user with dialogs.
5. **Optimistic concurrency.** Every write must send the hash from its last read (`expectedHash` or `If-Match`). A module the user edited in the meantime gives 412 rather than being overwritten.
6. **Backups.** The first time a module is changed, and every time one is replaced or deleted, its current version is exported to `%LOCALAPPDATA%\Rubberduck\ApiBackups`. If the backup fails, the change is not made.
7. **Wrong-token lockout.** Ten wrong tokens from one address lock it out for a minute.
8. **The Immediate window is a different class of power.** Anything entered in it is executed with the user's rights, and VBA can read, change and delete files and start programs. So: it is **off** unless the user ticks a separate setting ("Also let agents use the Immediate window"); with it on, the first command needs its own approval dialog, which says plainly that code will run; switching the setting or the API off, or regenerating the token, takes the approval back. The write approval does not cover it, and its approval does not cover writes. The MCP tools for it are only listed while it is allowed. Commands are limited to 50 lines and 10,000 characters per call, and to characters the editor's ANSI Immediate window can take, so a command is never silently altered.

Code edits go through the VBE's code module, as typing does, so attributes and a form's layout are untouched. Whole-module writes (`text`) go through the VBE's own export and import, as Git sync does, using a temp folder that exists for the length of the call.

Diagnostics only read: they ask the parser to catch up with the editor, wait for it, and run the inspections. They need no approval.

The Immediate window is read through UI Automation (the editor draws its own text, so it has no window text to read) and written by posting characters, then Enter as a key, to that one window. Nothing goes through the keyboard, so no other program can receive the keystrokes and the window does not need focus.

## Consequences

- `HttpListener` needs no new dependency and works on .NET Framework 4.6.2, but it cannot do TLS without certificate binding. Encrypted transport on the network is not offered in V1.
- The first approval covers all later writes, to any open project. Per-project approval was left out for simplicity.
- Everything touching the VBE runs on the UI thread and waits there, so a modal dialog or running code makes requests wait, and after 60 seconds fail with 503.
- A sandbox that can reach the host can read all open VBA code with the token. Treat the token like a password.
- With the Immediate window allowed, whoever holds the token and gets the user's one approval can run code. The approval dialog and the separate setting are the safeguards; there is no per-command confirmation.
- Running unit tests is treated like the Immediate window: same setting, same approval, MCP tool listed only while allowed. Tests are discovered by the engine only when the parser is ready, so listing waits for a parse. Results are remembered per test from the engine's events, since the engine keeps them nowhere readable; a test not run since Rubberduck started shows `Unknown`.
- Batch edits check everything they can before touching anything, but are not transactional: a failure partway leaves the earlier edits in place (backed up), and the error says so.
- Closing the editor's error dialogs relies on recognising them (a dialog of the editor's process, owned by its main window, with a Help button) and pressing End or OK by name; in a localised editor the names may differ, in which case the dialog is reported as blocking instead of being closed.
- Diagnostics runs share the inspector with the Inspection Results window, so a run is serialised with other API runs but can overlap one started from the UI.
- Reading and typing into the Immediate window depends on how the editor exposes it (UI Automation text, window messages). It was tried on the Office 365 editor; other versions may differ.
