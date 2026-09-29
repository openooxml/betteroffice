# DOCX: the worker as the only document owner

This note covers large documents in the DOCX editor's hybrid runtime, where a host engine sits on the main thread and a resident engine runs in the worker. It proposes an order for moving document ownership into the worker so that memory and open time improve step by step, before the phase 4 default flip in [editor-worker-convergence.md](editor-worker-convergence.md). The target architecture, failure rule, recovery log and API policy are the ones set out there; this note adds only the order and the measurements behind it.

## What the hybrid runtime costs today

Measured in Chromium on a private 1,209-page document, with main plus #909, #914, #901, #917 and #928:

| Stage | Where | Cost |
| --- | --- | --- |
| `openDocx`: parse, lower and seed | main | one 2.4 s task; the editing core's memory peaks at 747–881 MB and keeps 395 MB live |
| Worker snapshot (`residentWorkerSnapshot`, 95 MB of yrs state) | main | 0.6 s |
| `loadState`, fonts, a 16-page provisional layout and its frame | worker | 0.55 s |
| Full layout (`completeLayout`) | worker | 1.7 s |
| Display pages as they are built | worker | memory grows to 3.0 GB once every page is built |
| Ten host proposals, and the relayout they trigger | main | five tasks, up to 7.9 s |

The document is parsed and seeded twice: from bytes on the main thread, then from a snapshot in the worker. After an edit it is laid out twice: `layOutHere` on the main thread for geometry, and the worker for its frame. The main replica costs about 400 MB live and 750–880 MB of wasm memory that never shrinks. Before #946, a worker that ran out of memory moved display building to the main thread, which then reached 3.85 GB of its 4 GiB.

## Order

Each step ships behind the runtime flag, with the gates below, and removes one duplicate before the next begins.

1. **The worker opens the bytes.** The bytes are transferred to the worker, which parses, seeds, lays out the first pages and paints. The main thread gets the worker's state as one yrs update and applies it after the first paint, keeping its replica for the APIs that still read it. This takes the main-thread parse, the snapshot encode and the worker's `loadState` off the critical path; memory does not change yet. The main replica is a peer of the worker's, so later updates flow as they do today. Until the replica is ready, host calls await it, as `flushPendingInput` already does for pending input.
2. **One layout after an edit.** The main thread stops running `computeLayout` for edits. The layout reply the worker already sends for its frame becomes the main thread's geometry projection: page count, section and page rectangles, and note placement. This removes the second whole-document layout and the main thread's pagination memory.
3. **Reads come from the worker.** Display queries, `readParagraphs`, `findText`, revisions, content controls and `exportStructuredWithPages` are answered in the worker. The main thread keeps a small cache of query pages near the viewport, evicted by distance, instead of per-document query stores.
4. **Writes go through the worker, and the main replica goes.** Input, commands, edit batches, proposals, undo and redo, and collaboration updates are applied only in the worker, and save runs there. The main thread no longer instantiates the editing core, which saves its 750–880 MB. With no second engine, "no main-thread fallback" becomes structural.

The synchronous `DocxEditorRef` members keep working from projections: `getDocument` from the JS model, materialized lazily as today, and `getPositionAtPoint`, `scrollToParaId`, `scrollToCommentId` and `scrollToChangeId` from the last geometry projection. The convergence plan's async twins follow the same policy.

## Memory budgets

Once the worker is the only owner, its memory is the document's memory. `onMemoryPressure` and `getMemoryStats()` (#946) are the display side's signals for evicting built pages. `memoryBudget.workerLimitBytes` is the host's hard cap. A worker that runs out of memory is replaced once and then reported, and from step 4 nothing can take its place silently.

## Gates

Each step passes the convergence plan's phase gate for DOCX. It also holds these on a generated 1,000-page fixture with notes, floating tables and sections:

- bytes to first paint, and bytes to a complete layout;
- the main thread's and the worker's wasm memory after open, and after 10 host proposals;
- keystroke-to-paint p95 in the body and in a table cell;
- no `*_bg.wasm` samples on the main thread after open (from step 4).

## Risks

- **Typing latency.** Every keystroke already round-trips to the worker, which paints the caret itself. Steps 3 and 4 add no round trip on the typing path.
- **Collaboration.** Providers apply updates synchronously today. Step 4 needs the async replica described in the convergence plan.
- **Crashes.** With one owner, a crash loses unsaved edits. The recovery log and fresh-replica reopen from the convergence plan are prerequisites for step 4, not for steps 1–3.
- **Early host calls.** Step 1 changes when the main replica exists. A test lists the host APIs that read it, and covers each before and after the replica is ready.
