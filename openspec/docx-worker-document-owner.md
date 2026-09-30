# DOCX: the worker as the only document owner

DOCX is the first format to converge on the runtime in [editor-worker-convergence.md](editor-worker-convergence.md). This note is the DOCX slice order: each slice is one mergeable PR behind `experimentalWorkerOpen`, removes one reason the main thread still holds the document, and changes no default until the last. The target, failure rule, recovery log and API policy are the convergence plan's.

## Why

On a 1,209-page document, the hybrid runtime parses and seeds the document on the main thread (one 2.4 s task, 747–881 MB peak, 395 MB live), encodes a 95 MB snapshot for the worker (0.6 s), and lays out every non-resident edit twice. Its main replica keeps 750–880 MB of wasm memory that never shrinks. With one owner, the worker's memory is the document's memory and the main thread runs no wasm after `open`.

## In place or in review

- **Worker open.** The worker receives the bytes, parses, seeds, lays out and paints. The main replica is a yrs peer, hydrated from one worker update after the layout completes, with a context-only entry that attaches the retained source for save.
- **Worker layout.** The opened document, host edit batches, remote updates and revision-preview changes lay out in the worker alone. Local edits still lay out on the host.
- **On-demand replica.** A read-only, non-collaborating editor hydrates the replica only when an API, plugin, sidebar, pointer or key needs it. Hydration keeps the order of the calls waiting on it and moves neither the caret nor the scroll position.
- **Worker proposal authority.** The worker holds the proposal registry and the main thread mirrors it. Proposals, their geometry and paragraph navigation, `readParagraphs`, and the async `getParagraphIdentities` and `resolveParagraphAnchors` are served by the worker (`documentRead`), and plugins load once the worker is ready.

What still reads the replica: the synchronous `DocxEditorRef` members and plugin geometry reads, save and the comment projection, editing input, selection, hit testing and the caret, the sidebar's story projection, and collaboration.

## Slices

Each slice lands on main before the next depends on it.

- **(a) Async twins.** Every synchronous `DocxEditorRef` member that reads or writes the document, and the plugin geometry reads (`getPositionAtPoint`, `getAnchorGeometry`), gets an async twin; the synchronous member is deprecated in the same release. Reads are served by the worker through new `documentRead` kinds and never hydrate the replica. Mutations and selection-dependent members await the replica until (b) and (c) move them, which then changes no API. The replica-access table in `useDocxEditorRefApi.ts` classifies every member and fails on an unclassified one. Additive; no behavior change. Alongside it, CI imports every published web entry in plain Node without a DOM, and any top-level DOM access it finds is moved into effects or first calls.
- **(b) Save and comments in the worker.** Save runs in the worker session, including what `handleSave` and `writeEditorDocument` do on the main thread today: the host-owned comment projection (`withSavedComments`), reply and tracked-change range markers, and patching the last saved buffer. Comments, replies and resolution become session state that the worker saves and checkpoints. Render and save then read one authority. This follows the comment-marker fixes in the same save path (#1043).
- **(c) Input, selection and caret through the worker.** Compatibility input, commands, undo and redo apply in the worker, and local edits lay out there alone. Selection, hit testing, range rects, vertical moves and the selection context become worker queries. The main thread keeps an optimistic caret from the last geometry, and IME composition stays in the DOM and commits once. The worker already paints the caret for resident input; compatibility input needs the same latency budget. From here, `experimentalWorkerOpen` probes the environment at construction and throws instead of opening on the main thread.
- **(d) Async collaboration.** The DOCX provider talks to the worker through `AsyncCollaborationReplica`: async state vectors and diffs, incoming updates applied in order and awaited, origins preserved, the connection generation checked after every await, and a peer marked synced only after its update applies. After (d) nothing hydrates the replica.
- **(e) Recovery.** The recovery log (a worker checkpoint at `ready` plus every accepted update, with the comment state and the proposal registry in the checkpoint) and the fresh-replica reopen, with the kill-the-worker spec: kill, reopen, keep typing, converge with a peer and save, including an unsaved comment reply and resolution.
- **(f) Default flip.** Breaking, one release after (a)'s deprecations. The worker path becomes the only one: the host engine, `residentWorkerSnapshot`, the replica, the automatic fallback (`fallback`, `replayInputOnMainThread`, `adoptHostEngine`), the host-image gate (`displayListNeedsHostImages`) and the deprecated synchronous members are deleted, every editor probes at construction, and `experimentalWorkerOpen` becomes a deprecated no-op.

(b) and (c) touch different code and can proceed in parallel after (a); (d) needs (c); (e) needs (b); (f) needs all of them.

## Memory

`onMemoryPressure` and `getMemoryStats()` are the display side's signals for evicting built pages, and `memoryBudget.workerLimitBytes` is the host's hard cap. The hybrid runtime replaces a worker that runs out of memory once and reports it; worker-held proposals already fail instead. From (f), running out of memory ends the session visibly, and a replacement comes only from an explicit reopen from the recovery log.

## Gates

Each slice passes the convergence plan's phase gate for DOCX and holds these on a generated 1,000-page fixture with notes, floating tables and sections, measured with interleaved arms at n≥5:

- bytes to first paint, and bytes to a complete layout;
- the main thread's and the worker's wasm memory after open, and after 10 host proposals;
- keystroke-to-paint p95 in the body and in a table cell;
- main-thread long tasks during open and per interaction;
- no `*_bg.wasm` samples on the main thread after open (from (d)).
