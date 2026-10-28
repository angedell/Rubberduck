# Glossary

**Git sync** — Compares a VBA project with a folder, once, when the user asks, and syncs what they pick. See `docs/GitSyncCompare.md`.

**API** — The HTTP server inside the add-in that lets AI agents read and edit modules and forms. Off by default. See `docs/RestApi.md` and `docs/adr/0001-local-api-for-ai-agents.md`.

**Agent** — A program, usually driven by an AI model, that talks to the API or the MCP server.

**Token** — The secret an agent sends as `Authorization: Bearer <token>`. Kept per user in `%APPDATA%\Rubberduck\api-token.txt`.

**Write approval** — The user's one-time yes to the first write after the API starts. Taken back when the API is switched off or the token regenerated.

**Hash** (of a module) — Identifies a module's code as of a read, ignoring line-ending differences. Every write sends it back; a mismatch is a 412.

**Code** vs **text** (of a module) — *Code* is what the code pane shows. *Text* is the module as the VBE exports it, with its hidden header (attributes, a form's layout). Editing code leaves the header alone; replacing text replaces both.

**Diagnostics** — The results of Rubberduck's code inspections for a project or module, as agents can fetch them: inspection, severity, description, position. Read-only.

**Running code** — Entering a command in the Immediate window, or running unit tests, executes the user's VBA. It is one separate opt-in ("Also let agents run code") with its own one-time approval, apart from write approval. Reading the Immediate window and listing tests do not run anything.

**Check** — The agent's "is my work good" call: parser errors, then inspections, then optionally tests, with one verdict (`clean`).

**Batch** — Several module edits in one request, one per module, every hash checked before the first is applied, one approval.

**Error dialog** — The editor's own run-time or compile error dialog. While an agent's Immediate command runs, it is closed with End and reported; any other dialog is left to the user and reported as `blockedByDialog`.

**Document module** — The code behind a sheet or the workbook. It cannot be removed, only emptied.

**Resources** — The ribbon XML (`customUI.xml`, `customUI14.xml`) and the pictures stored inside a document's file, shown as a node next to the project's references in the Code Explorer and offered to agents by the API. Only saved Open XML documents (xlsm, docm, pptm...) have them. See `docs/adr/0002-stage-resource-changes-until-the-document-closes.md`.

**Staged change** — A change to a resource that is kept aside, and shown as if made, until the document is closed and its file can be written.
