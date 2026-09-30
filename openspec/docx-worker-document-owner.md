# DOCX: the worker as the only document owner

DOCX is the first format to converge on the runtime in [editor-worker-convergence.md](editor-worker-convergence.md). This note is the DOCX slice order: each slice is one mergeable PR behind `experimentalWorkerOpen`, removes one reason the main thread still holds the document, and changes no default until the flip. The target, failure rule, recovery log and API policy are the convergence plan's.

## Why

On a 1,209-page document, the hybrid runtime parses and seeds the document on the main thread (one 2.4 s task, 747–881 MB peak, 395 MB live), encodes a 95 MB snapshot for the worker (0.6 s), and lays out every non-resident edit twice. Its main replica keeps 750–880 MB of wasm memory that never shrinks. With one owner, the worker's memory is the document's memory and the main thread runs no wasm after `open`.

## Prerequisites

- **Worker open.** The worker receives the bytes, parses, seeds, lays out and paints. The main replica is a yrs peer, hydrated from one worker update, with a context-only entry that attaches the retained source for save.
- **Worker layout.** The opened document, host edit batches, remote updates and revision-preview changes lay out in the worker alone. Local edits still lay out on the host.
- **On-demand replica.** A read-only, non-collaborating editor hydrates the replica only when an API, plugin, sidebar, pointer or key needs it. Hydration keeps the order of the calls waiting on it and moves neither the caret nor the scroll position.
- **Worker proposal authority.** The worker holds the proposal registry and the main thread mirrors it, with its geometry. While the replica is unhydrated, proposals, `readParagraphs`, `getParagraphIdentities`, `resolveParagraphAnchors` and paragraph navigation are served by the worker through `documentRead`; the first hydration hands the registry over and maps the worker's versions onto the replica's.

What still reads the replica: the synchronous `DocxEditorRef` members and plugin geometry reads, save and the comment state, the revision, comment and outline sidebars, editing input, selection, hit testing and the caret, and collaboration. Images still decode on the main thread, and a document with images presents from the main thread (`displayListNeedsHostImages`).

## Slices

- **(a) Async twins.** Needs the worker proposal authority. Every synchronous member in the convergence plan's inventory gets an async twin, and the synchronous member is deprecated in the same release. Reads the worker can answer are served by new `documentRead` kinds while it holds the document, and never hydrate the replica; the others await the replica until a later slice moves them, which then changes no API. A twin that returns a version goes where the edits consuming it run, as `readParagraphs` does. Synchronous reads of pushed projections stay: `getAnchorGeometry` for proposal targets reads the mirror, and only its other targets are deprecated in favor of `readAnchorGeometry`. Compile-time tables classify every member of the ref, the props, and the plugin context and geometry. Additive; no behavior change.
- **(b) Save and comment state in the worker.** After the comment-marker fixes in the same save path land. Save runs in the worker session, including what `handleSave` and `writeEditorDocument` do on the main thread today: the host-owned comment projection (`withSavedComments`), reply and tracked-change range markers, and patching the last saved buffer. Comments, replies and resolution become session state that the worker saves and journals. Until the flip, the deprecated `replyToComment`, the synchronous `resolveComment` and a controlled `comments` prop reach it through an ordered comment-state bridge; `getComments` reads the comment list the worker pushes. Until (c2), local edits still apply on the replica first. Save awaits both, so the worker has every edit and comment change before it reads.
- **(c1) Display queries in the worker.** Hit testing, range rects and vertical moves become worker queries, since they derive from the display and not from the edit authority. Each answers against the presented layout: a frame the worker laid out is answered there, and until (c2) a frame laid out on the host after a local edit is answered on the host.
- **(c2) Writes, selection and versioned reads in the worker.** These move together: compatibility input, commands, undo and redo, `applyEdits` and `validateEdits`, `findText`, content controls, `exportStructuredWithPages`, and selection. Local edits lay out in the worker alone, and from here render and save read one authority. A worker write resolves only after a hydrated replica has applied its update. The main thread keeps an optimistic caret from the last geometry, and IME composition stays in the DOM and commits once; compatibility input gets the latency budget resident input has. From here, `experimentalWorkerOpen` runs the capability probe at construction and throws instead of opening on the main thread.
- **(d) Async collaboration.** Needs (c2). The DOCX provider talks to the worker through `AsyncCollaborationReplica`: async state vectors and diffs, incoming updates applied in order and awaited, origins preserved, the connection generation checked after every await, and a peer marked synced only after its update applies.
- **(e) Recovery.** Needs (b), (c2) and (d). On the `experimentalWorkerOpen` path, a worker failure no longer hands the session to the main thread: the automatic takeover (`fallback`, `replayInputOnMainThread`, `adoptHostEngine`) is replaced by the explicit reopen from the recovery log (a worker checkpoint at `ready`, then every accepted update, comment change and proposal registry change in one ordered journal) into a fresh replica. From (e), a deprecated synchronous member on an unhydrated replica throws `DocxReplicaNotReadyError` on this path even when the worker holds nothing the source lacks, since hydrating by reopening the source is the takeover (e) removes; the default path keeps it until (f). The kill-the-worker spec kills the worker, reopens, keeps typing, converges with a peer and saves, including an unsaved comment reply, an undecided proposal that keeps its retry identity, and a decision on another.
- **(g) Images in the worker.** Independent of the other slices. Package media decode in the worker with `createImageBitmap`, SVG included, and the worker presents documents with images, with an image-parity check against the main-thread paint.
- **(h) Built-in sidebars from worker projections.** Needs (b) for comments. The revision, comment and outline sidebars read versioned worker projections (`listRevisions`, comment anchors, headings) and navigate through worker queries, including off-window targets.
- **(f) Default flip.** Breaking, and at least one release after the last deprecation. The host engine, `residentWorkerSnapshot`, the replica, the automatic takeover, the host-image gate and the deprecated synchronous members are deleted; every editor runs the probe at construction; `experimentalWorkerOpen` becomes a deprecated no-op. Needs every slice above.

The plain-Node import check for every published package is independent of the slices. Every slice except (a) is measured with phase 0's harness and waits for it; (c2) also needs phase 0's capability probe.

## Memory

`onMemoryPressure` and `getMemoryStats()` are the display side's signals for evicting built pages, and `memoryBudget.workerLimitBytes` is the host's hard cap. The hybrid runtime replaces a worker that runs out of memory once and reports it, except while the worker holds proposals the replica lacks. From (e), running out of memory ends the session on the opt-in path, and a replacement comes only from an explicit reopen from the recovery log.

## Gates

Each slice passes the convergence plan's phase gate for DOCX with the worker checks it implements (the construction error from (c2), kill-the-worker from (e), everything from (f)), and every slice after (a) holds these on a generated 1,000-page fixture with notes, floating tables and sections, measured with interleaved arms at n≥5:

- bytes to first paint, and bytes to a complete layout;
- the main thread's and the worker's wasm memory after open, and after 10 host proposals;
- keystroke-to-paint p95 in the body and in a table cell;
- main-thread long tasks during open and per interaction;
- no `*_bg.wasm` samples on the main thread after open, once (b), (c2), (d), (g) and (h) have landed, for a host that uses no deprecated synchronous member; for every host from (f).
