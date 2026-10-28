# 0002. Ribbon XML and pictures are staged until the document is closed

## Status

Accepted.

## Context

The Code Explorer's **Resources** node (and the API) shows the ribbon XML (`customUI/customUI.xml`, `customUI14.xml`) and the pictures (`customUI/images`) that live inside an Open XML document's file (xlsm, docm, pptm...). They are not part of the VBA project, so the VBE cannot read or change them. The only way is the file itself, which is a zip of parts.

Office keeps that file locked for as long as the document is open: it can be read, not written.

## Decision

Changes are **staged**, not written. A listing or read already includes them, so the user and agents see their own work at once. Staged changes are kept per document in `%LOCALAPPDATA%\Rubberduck\PendingResources`, so they survive Rubberduck restarting.

They are written into the file when it is free: when its project is closed (retrying for two minutes), when Rubberduck starts, and as a last try when it shuts down. The writer rebuilds the zip next to the original, keeps the original in `%LOCALAPPDATA%\Rubberduck\ResourceBackups`, and swaps it in with `File.Replace`, so a failure leaves the document as it was. A document that is still open is never forced: its changes simply stay staged.

Rejected: closing, patching and reopening the document on every edit. It would kill the VBE project and Rubberduck's state mid-session, and fails outright when the add-in lives in that very file.

## Consequences

- Nothing a user or agent stages reaches the document file until Office lets go of it. The tree marks pending items with `*`, and **Discard Pending Resource Changes** throws them away.
- If Office quits before Rubberduck gets its chance (Rubberduck unloads first), the changes wait for the next time Rubberduck runs with the document closed.
- Only the two fixed ribbon part names and the `customUI/images` folder are handled. A document whose ribbon part lives elsewhere is shown without it.
- The project's `rubberduck.json` is kept in the document too, as `customUI/rubberduck.json` with its own root relationship. It is created, staged like any other change, when the tree first shows a document without one; discarding it does not bring it back until the next session. It survives Excel saving the workbook (checked against a real Excel). Word and PowerPoint are not verified.
- Git sync treats the ribbon XML (`Resources/customUI*.xml`), pictures (`Resources/images/*`) and `rubberduck.json` as items like modules; importing stages them.
- Agent writes need the same one-time write approval as module writes, and the same hash check against stale reads.
