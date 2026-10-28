# Changes in this fork

This fork of [Rubberduck](https://github.com/rubberduck-vba/Rubberduck) adds **Git Sync**, **bulk removal** to the Code Explorer, and an **API for AI agents**. Everything below is on top of the original; the rest of this README is unchanged.

## API for AI agents
A local HTTP server and MCP server that lets AI agents read and edit the modules, classes and forms of the open projects without exporting anything to a folder. Off by default (Settings > General); needs a token; the first write asks you once; a toggle lets agents in a sandbox or VM connect over the network. It can also run Rubberduck's code inspections for the agent (diagnostics) and, behind a separate opt-in and approval, use the Immediate window and run the unit tests, which runs code. It also answers questions about the code (parser errors, symbols, references, libraries, metrics) and can apply a batch of edits at once. It never touches cells. Guide: [docs/RestApi.md](docs/RestApi.md). Design and security model: [docs/adr/0001-local-api-for-ai-agents.md](docs/adr/0001-local-api-for-ai-agents.md).

## Git Sync
<img width="322" height="87" alt="image" src="https://github.com/user-attachments/assets/2d82a89e-e56f-4d74-9826-aa269dac9783" />

Compares a VBA project with a folder on disk, subfolders included, so the code can live in a git repository. It runs once, when you ask: nothing is watched or polled.

- **Where:** Code Explorer, right-click a project, **Sync Project > Git Sync...**, or the toolbar's **Git sync** button. The folder is asked for the first time and remembered per workbook; **Change folder...** in the review picks another.
- **Review:** both sides are scanned and what is out of sync is listed: only in the project, only in the folder, changed on one side, changed on both, or different with no history. Tick what to sync, use **Select all / none**, and for items where the history cannot tell which side wins use **Use project version** or **Use folder version**. Nothing happens until **Sync selected**.
- **Compare and merge:** select an item that exists on both sides and press **Compare...** (or double-click it) for a three-pane view in the manner of WinMerge: **Project**, **Result** and **Folder**. What is the same on all three is black; what differs is colored, with the changed words marked inside changed lines. The pointer turns into a hand over a differing block: **click** it to put it in the Result, click again to take it back, **Ctrl+click** to take both. **All from project / All from folder** take a whole side, and you can still cherry pick on top. The Result is editable by hand. **Merge** makes the item a *Merge*: on **Sync selected** the project and the file both get the Result (the old versions are backed up). Indentation and trailing space are not shown as differences. The comparison is a C# port of WinMerge's own engine (the diffutils line diff and the word-level diff of changed lines; GPL-2.0-or-later, see the headers of `WinMergeLineDiff.cs` and `WinMergeWordDiff.cs`). Step-by-step guide: [docs/GitSyncCompare.md](docs/GitSyncCompare.md).
- **Ignore for good:** **Ignore permanently** adds the selected items to `sync.ignore` in `rubberduck.json`. Ignored items are not listed again.
- **Through the VBE:** modules are written with the VBE's own export and brought in with `VBComponents.Import`, so hidden attributes survive. Document modules (`ThisWorkbook`, sheets) cannot be removed, so only their code is replaced.
- **Normalized comparison:** both sides are compared after indenting the code with the Rubberduck indenter, so line endings, indentation and trailing space never count as a change. Modules with `@NoIndent` are compared as they are.
- **Line endings:** the VBE reads a corrected copy with CRLF endings; the file in the repository is not rewritten. If git converts line endings (`core.autocrlf`, `.gitattributes`) that is fine.
- **Folders:** a file's location below the sync folder sets its `@Folder` annotation (`Models\Orders\Foo.cls` becomes `'@Folder("Models.Orders")`). A module moved to another virtual folder moves its file. Files directly in the sync folder keep the `@Folder` they have.
- **History:** `rubberduck.sync.json`, next to `rubberduck.json` (or in the sync folder if there is none), holds the normalized signature of each module and file as of the last sync. Commit it: comparing today's signatures with it tells which side changed, on any machine.
- **Backups:** before a module or file is overwritten, the old version is saved in `%LocalAppData%\Rubberduck\GitSyncBackups\<time>`. If an import fails after the old module was removed, it is put back.
- **`rubberduck.json`:** found in the sync folder or any parent, like `biome.json`. Keys that are left out keep your own settings; unknown keys are reported.
  ```json
  {
    "indenter": { "indentSpaces": 2, "indentCase": true },
    "sync": {
      "ignore": ["/modScratch.bas", "Sandbox/"],
      "enforce": ["AttributeValueOutOfSync", "MissingAttribute", "MissingModuleAnnotation", "ModuleWithoutFolder"],
      "onViolation": "fix",
      "ignoreCaseOutsideStrings": true
    }
  }
  ```
  `indenter` takes the names of the indenter settings and also applies to the editor while the project is the active one. `enforce` lists the rules a module or file has to follow to be used as the source of a sync; with `"fix"` (the default) the sync repairs it as it goes, the way Rubberduck's own quick fixes would (a `True` attribute gets its `'@PredeclaredId`/`'@Exposed` annotation, an annotation sets its attribute, `VB_Name` is added or corrected, a missing `@Folder` is set from the file's folder) and writes the repaired text to both the module and the file; with `"block"` it cannot be ticked until fixed by hand, with `"warn"` it is only flagged. A module with no `@Folder` whose file is in the root cannot be repaired automatically.
- **Case:** VBA does not care about upper and lower case, so by default (`"ignoreCaseOutsideStrings": true` in `rubberduck.json`) a change of case is not a change, except inside a string literal, where it counts. Comments count as code here. Set it to `false` to treat every change of case as a change. Changing the setting changes how modules compare, so modules that differ may show up as changed on both sides once until they are synced again.
- **Excluding things:** besides `sync.ignore`, a `.rdsyncignore` file in the sync folder or any subfolder works, using the `.gitignore` syntax. It is a separate file so it cannot be confused with git's own `.gitignore`.
- **Never deletes:** nothing is ever deleted from the project, and deleting a file does not delete its module, so a git checkout cannot wipe modules.

## Select mode and bulk removal
<img width="778" height="717" alt="image" src="https://github.com/user-attachments/assets/9a15350e-d46b-4716-ba34-d2b10e4520c2" />

- The Code Explorer toolbar has a **select** toggle that shows a checkbox next to every module and folder (ticking a folder ticks what is in it). Projects, members, references and document modules cannot be ticked.
- While it is on, the red button removes all ticked modules at once. One dialog asks whether to export first, whether to keep the folder structure or use a single folder, and where to (the last folder is remembered). If the export fails, nothing is removed. Select mode switches off afterwards.

---

<!-- ![banner](https://user-images.githubusercontent.com/5751684/113501222-8edfe880-94f1-11eb-99a9-64583e413ef3.png) -->

## Links

- [**Installing**](https://github.com/rubberduck-vba/Rubberduck/wiki/Installing)
- [Contributing](https://github.com/rubberduck-vba/Rubberduck/blob/next/CONTRIBUTING.md)
- [Attributions](https://github.com/rubberduck-vba/Rubberduck/blob/next/docs/Attributions.md)
- [Wiki](https://github.com/rubberduck-vba/Rubberduck/wiki)
- [Website](https://rubberduckvba.com)
- [Blog](https://rubberduckvba.blog)
- [Shop](https://ko-fi.com/rubberduckvba/shop)

<a href='https://ko-fi.com/N4N2IWEIG' target='_blank'><img height='36' style='border:0px;height:36px;' src='https://storage.ko-fi.com/cdn/kofi1.png?v=3' border='0' alt='Support us on ko-fi.com' /></a>

## Releases

- The [latest release](https://github.com/rubberduck-vba/Rubberduck/releases/latest)
- See [all releases](https://github.com/rubberduck-vba/Rubberduck/releases) including pre-release tags

---

## [License (GPLv3)](https://github.com/rubberduck-vba/Rubberduck/blob/next/LICENSE)

Copyright &copy; 2014-2023 Rubberduck project contributors.

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the [GNU General Public License](https://www.gnu.org/licenses/gpl-3.0.en.html) for more details.

---

## What is Rubberduck?

The Visual Basic Editor (VBE) has stood still for over 20 years, and there is no chance a first-party update to the legacy IDE ever brings it up to speed with modern-day tooling. Rubberduck aims to bring the VBE into this century by doing exactly that.

Read more about contributing here:

[![contribute!](https://user-images.githubusercontent.com/5751684/113513709-071dcc80-9539-11eb-833d-d21532065306.png)](https://github.com/rubberduck-vba/Rubberduck/blob/next/CONTRIBUTING.md)

The add-in has *many* features - below is a quick overview. See https://rubberduckvba.com/features for more details.

### Enhanced Navigation

The Rubberduck *command bar* displays docstring for the current member.

![command bar](https://user-images.githubusercontent.com/5751684/113501975-25fb6f00-94f7-11eb-9189-fcf2a0dd98da.png)

The *Code Explorer* drills down to member level, has a search bar, and lets you visualize your project as a virtual folder hierarchy organized just the way you need it.

All references to any identifier, whether defined in your project or any of its library references, are one click away. If it has a name, it can be navigated to.

### Static Code Analysis, Refactorings

Rubberduck analyses your code in various configurable ways and can help avoid beginner mistakes, keep a consistent programming style, and find all sorts of potential bugs and problems. Many code inspections were implemented due to frequently-asked [VBA questions on Stack Overflow](https://stackoverflow.com/questions/tagged/vba), and on many occasions, an automatic quick-fix is available.

Rename variables to meaningful identifiers without worrying about breaking something. Promote local variables to parameters, extract interfaces and methods from a selection, encapsulate fields into properties, reorder and/or delete parameters, and automatically update all callers.

### Unit Testing

Write code that *provably* works, by invoking it from small test procedures that setup the conditions for a test case and assert that the expected outcome happened. Rubberduck provides a rich MSTest-inspired API, and soon an experimental mocking framework (a COM-visible wrapper around [Moq](https://github.com/Moq)) that can automatically implement VBA interfaces and configure mock objects.

![test explorer](https://user-images.githubusercontent.com/5751684/113502368-fa2db880-94f9-11eb-954f-5735c15d4c3e.png)

### Smart Indenter

A port of the popular 32-bit add-in by Office Automation Ltd., whose legacy VB6 source code was generously made [freely available](https://github.com/rubberduck-vba/Rubberduck/tree/next/Rubberduck.SmartIndenter/Legacy) for Rubberduck under GPLv3 by the legendary Stephen Bullen and Rob Bovey themselves! Rubberduck will prompt to import your *Smart Indenter* settings on first load if detected.

### Annotations

Special comments that become a game changer with Rubberduck processing them: organize modules in your project using `@Folder` annotations, synchronize `VB_Description` and `VB_PredeclaredId` hidden attributes without manually exporting, editing, and re-importing modules with `@Description` and `@PredeclaredId` annotations.

### More?

Of course there's more! There's tooling to help synchronizing the project with files in a folder (useful for source/version control!), some auto-completion features like self-closing parentheses and quotes; there's a regular expression assistant, a replacement for the VBE's *add/remove references* dialog, and so many other things to discover, and yet even more to implement.

---

## Tips

Rubberduck isn't a lightweight add-in and consumes a large amount of memory. So much, that working with a very large project could be problematic with a 32-bit host, and sometimes even with a 64-bit host. Here are a few tips to get the best out of your ducky.

- **Start small**: explore the features with a small test project first.
- **Refresh often**: refresh Rubberduck's parser every time you pause to read after modifying a module. The fewer files modified since the last parse, the faster the next parse.
- **Review inspection settings**: there are *many* inspections, and some of them may produce *a lot* of results if they're directly targeting something that's part of your coding style. Spawning tens of thousands of inspection results can significantly degrade performance.
- **Avoid late binding**: Rubberduck cannot resolve identifier references and thus cannot understand the code as well if even the VBA compiler is deferring all validations to run-time. Declare and return explicit data types, and cast from `Object` and `Variant` to a more specific type whenever possible.

Join us on our [Discord server](https://discord.gg/MYX9RECenJ) for support, questions, contributions, or just to come and say hi!

For more information please see [Getting Started](https://github.com/rubberduck-vba/Rubberduck/blob/next/docs/GettingStarted.md) in the project's wiki, and follow the project's blog for project updates and advanced VBA OOP reading material.

---

## Roadmap

After over two years without an "official" new release, Rubberduck version jumped from 2.5.2 to 2.5.9, adding minor but interesting features to an already impressive array.

### The road ahead

Rubberduck 2.x is now planned to end at 2.5.9.x, perhaps with a number of small revisions and bug fixes, but nothing major should be expected, as the developers' attention is shifting to the 3.0 project:

- Parsing and understanding VBA code is moving to a language (LSP) server
- We're making a new editor _inside_ (for now) the Visual Basic Editor that will be the LSP client
- Baseline server-side feature set for 3.0 is everything 2.5.9 does
- Baseline client-side feature set for 3.0 is the 2.5.x UI (perhaps tweaked a bit/lot) hosted in the Rubberduck Editor

Fully controlling the editor opens Rubberduck to everything we ever dreamed of:

- In-editor syntax and static code analysis reporting and quick-fixing
- Full editor theming, custom syntax highlighting

See the [Rubberduck3](https://github.com/rubberduck-vba/Rubberduck3) repository for more information.
