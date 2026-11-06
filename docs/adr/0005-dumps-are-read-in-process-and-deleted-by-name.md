# 0005. Dumps are read in process and deleted by name

## Status

Proposed.

## Context

A memory dump (`POST /api/v1/dump`, ProcDump) is the one way to see why the editor hangs: the log and the API cannot say what a stuck thread is doing. Until now the dump was left in a folder for a person to open in a debugger. Three things went wrong in the first real use: the dump was written on the machine that runs Excel, not the one the agent works on, so it could not be read; the folder was one Dropbox syncs, so a 2 GB image of Excel's memory, open workbooks included, was on its way to the cloud; and nothing could delete it.

## Decision

1. **The editor reads its own dumps.** `IDumpAnalyzer` (ClrMD, `Microsoft.Diagnostics.Runtime`) opens a full dump in the editor's process and lists the managed threads with their stacks. A dump has to be read by the runtime that wrote it (the debugging library `mscordacwks`/`mscordaccore` must match), and that runtime is the one in the same process. The answer is findings, not a thread dump: `DumpDiagnosis` turns threads into what is blocked, on what, and who is in the way (a main thread waiting while another thread calls into the host is a probable deadlock; inspections still running are named). It knows nothing of ClrMD and is tested with made-up stacks.
2. **Only managed threads.** A hang in native code (a dialog, an OLE call inside Excel) is reported as a main thread without a cause, and the summary says so. Reading native stacks would need a debugger engine and symbols; not worth it until a hang needs it.
3. **A monitor, because hangs do not wait to be asked.** `POST /api/v1/dump/monitor` starts ProcDump with its own triggers (hung window, unhandled exception, processor use) and a dump limit. `GET .../monitor?analyze=true` reads the newest dump it wrote. The monitor needs the approval a dump needs, since it takes dumps by itself.
4. **Deletion is part of the feature.** `DELETE /api/v1/dump/{name}` and `DELETE /api/v1/dump?keep=n`, and `?delete=true` on the analysis. They need the approval writes need. A dump is addressed by its file name alone, resolved against the files the dump folder holds: a path, a directory part or another extension is refused, so the route cannot delete anything else.
5. **A synced folder is called out.** `GET /dump` and the capture answer carry a `warning` when the dump folder's path looks like Dropbox, OneDrive, Google Drive or the like.
6. **Reading needs no approval.** It sends back stacks and exception messages, not memory, and takes nothing from the editor. Capturing, monitoring and deleting keep the approval.

Rejected: shipping the dump to the agent's machine (it is the editor's whole memory; reading it where it is written keeps it from travelling); running `dotnet-dump` or WinDbg as a child process (not installed on the machines that matter, and the output would have to be parsed); a separate analysis process (more robust against a crash of the reader, but a second thing to ship and start; revisit if reading a large dump ever takes the editor down).

## Consequences

- Rubberduck.Core takes a dependency on ClrMD and, through it, on newer `System.Collections.Immutable`, `System.Memory`, `System.Buffers`, `System.Runtime.CompilerServices.Unsafe` and `System.Threading.Tasks.Extensions`. The add-in resolves assemblies from its own folder by name, so versions do not clash inside Excel; the test project needed binding redirects (`RubberduckTests/app.config`).
- Reading a large dump uses the editor's memory and a thread of the API for up to three minutes, in a process that is, by assumption, already in trouble. The limit is a cancellation checked between threads, not a hard stop.
- The analysis is only as good as its patterns (`DumpDiagnosis`): a hang with another shape shows as "nothing recognizable" until a pattern is added.
