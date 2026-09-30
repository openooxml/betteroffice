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

The document is parsed and seeded twice: from bytes on the main thread, then from a snapshot in the worker. Resident input is laid out in the worker alone, since `finishResidentMutation` skips the host relayout. Every other edit is laid out twice, by `layOutHere` on the main thread for geometry and by the worker for its frame. That covers compatibility input (selection replacement, suggesting mode, stored formatting and non-ASCII text), commands, edit batches and host proposals. The main replica costs about 400 MB live and 750–880 MB of wasm memory that never shrinks. Before #946, a worker that ran out of memory moved display building to the main thread, which then reached 3.85 GB of its 4 GiB.

## Order

Each step ships behind the runtime flag, with the gates below, and removes one duplicate before the next begins.

1. **The worker opens the bytes.** The bytes are transferred to the worker, which parses, seeds, lays out the first pages and paints. After the first paint, the main thread applies the worker's state as one yrs update, and keeps its replica for the APIs that still read it. A yrs update does not carry the package context that `openDocx` sets up: the retained source bytes, their digest and the source metadata that `materializeDocx` and save read. So the host also gets a context-only entry point that attaches the retained bytes without seeding.

   This takes the main-thread parse, the snapshot encode and the worker's `loadState` off the critical path; memory does not change yet. The main replica is a peer of the worker's, so later updates flow as they do today. Until the replica and its context are ready, async host calls await them, as `flushPendingInput` already does for pending input.
2. **One layout after an edit.** Worker edit replies do not carry a layout today; only `bootstrap`, `sync` and `completeLayout` do. So the prerequisite is edit replies that carry the retained layout projection, tied to the frame epoch: page count, section and page rectangles, and note placement. The main thread then stops running `computeLayout` for compatibility edits, commands and edit batches. One exception stays until step 3: `exportStructuredWithPages` reads the host's retained pagination, and relays out on the host when that is missing (`relayout({ onHost: true })`). Its version feeds host edit batches, so it moves with them. This removes the second whole-document layout, and the host's pagination memory except during a paged export.
3. **Versioned reads and writes move together.** Everything that returns or checks a version runs in the worker: `readParagraphs`, `findText`, `validateEdits`, `applyEdits`, content controls, proposals, `exportStructuredWithPages`, and `getPositionAtPoint`, whose result carries a version for edit batches. Versions are scoped to a replica, so a version read in one replica cannot be passed to an edit in the other. Display queries and revisions move with them.

   Undo stays on the host until step 4, so the relay that applies worker updates to the host must keep each change's history policy. Today every worker update is applied as local history (`applyLocalUpdate`). Proposals and `history: 'none'` batches must stay out of undo, and `history: 'separate'` batches must keep their boundaries. The gate adds typing, then a proposal, then a separate batch, then undo and redo.

   The proposal registry lives outside yrs: proposal ids, retry keys, decisions and `previewVersion`. A replacement worker must not lose it, so step 3 carries the registry in the worker's snapshot and rehydrates it on replacement, covered by a replacement test. Until that exists, replacing the worker fails the session visibly instead. The main thread keeps a small cache of query pages near the viewport, evicted by distance, instead of per-document query stores.
4. **The remaining writes go through the worker, and the main replica goes.** Compatibility input, commands, undo and redo, and collaboration updates are applied only in the worker, and save runs there. The main thread no longer instantiates the editing core, which saves its 750–880 MB. With no second engine, "no main-thread fallback" becomes structural.

Synchronous `DocxEditorRef` members follow the convergence plan's policy. That covers `getDocument`, `getPositionAtPoint`, `getSelectionInfo`, `getPageContent`, `findInDocument`, the navigation members that resolve live anchors and move the selection, and mutations that return results. Steps 1–3 keep them working in the hybrid runtime, where the main replica still exists. Each gets an async twin, a test lists them, and they are deprecated one release before step 4, after which they stay available only under `'inline'`.

## Memory budgets

Once the worker is the only owner, its memory is the document's memory. `onMemoryPressure` and `getMemoryStats()` (#946) are the display side's signals for evicting built pages. `memoryBudget.workerLimitBytes` is the host's hard cap. In the hybrid runtime, #946 replaces a worker that runs out of memory once and then reports it. From step 4 the convergence plan's failure rule applies instead: running out of memory ends the session visibly, and a replacement comes only from an explicit reopen from the recovery log.

## Gates

Each step passes the convergence plan's phase gate for DOCX. It also holds these on a generated 1,000-page fixture with notes, floating tables and sections:

- bytes to first paint, and bytes to a complete layout;
- the main thread's and the worker's wasm memory after open, and after 10 host proposals;
- keystroke-to-paint p95 in the body and in a table cell;
- no `*_bg.wasm` samples on the main thread after open (from step 4).

## Risks

- **Typing latency.** Resident input already round-trips to the worker, which paints the caret itself whenever the painted-caret mode allows. Step 4 moves compatibility input to the worker too, and needs the same latency budget for it.
- **Collaboration.** Providers apply updates synchronously today. Step 4 needs the async replica described in the convergence plan.
- **Crashes.** With one owner, a crash loses unsaved edits. The recovery log and fresh-replica reopen from the convergence plan are prerequisites for step 4, not for steps 1–3.
- **Early host calls.** Step 1 changes when the main replica exists. A test lists the host APIs that read it, and covers each before and after the replica is ready.
