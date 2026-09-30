# Editor worker convergence plan

This plan converges the DOCX, PPTX and XLSX web editors on one runtime, per [#535](https://github.com/openooxml/betteroffice/issues/535). A dedicated worker owns the only writable document handle from `open` to `dispose` and paints into transferred `OffscreenCanvas` surfaces. The main thread keeps React, focus, keyboard/IME, pointer input, overlays and the accessible projection. DOCX converges first; PPTX and XLSX follow on the same contract. The formats share the runtime boundary, not the document model. Headless, native and Python APIs are unaffected, and their signatures do not change.

## Decisions

- **One authority.** The worker holds the wasm handles, undo, layout and recalculation caches, fonts, decoded images and save. The main thread holds no engine, no replica and no display list. It receives small projections: state, version, dirty flag, surface sizes, caret/selection/hit geometry, query results, change events and yrs updates.
- **One render path.** Layout and paint share the worker, so display lists never cross the boundary. Thumbnails and exports come back as transferred `ImageBitmap`s or bytes.
- **No fallback.** There is no main-thread runtime and no automatic takeover. A worker failure ends the session visibly; recovery is an explicit reopen from the recovery log.
- **Fail loud at construction.** The editor probes once, before it opens anything, for `Worker`, `OffscreenCanvas` with a 2D context, `HTMLCanvasElement.transferControlToOffscreen` and `createImageBitmap`. When one is missing, construction throws a typed `UnsupportedEnvironmentError` that names the missing capability, and the editor renders that error. Safari 16.4, the desktop app's target (`apps/desktop/vite.config.ts`), has all of them.
- **Transfer and `postMessage` only.** Bytes, fonts, canvases and bitmaps move as transferables. There is no `SharedArrayBuffer`, so hosts need no COOP/COEP headers.
- **Import-safe under SSR.** Importing any web package, including the React editors, touches no `window`, `document`, `self`, `navigator`, `Worker`, `OffscreenCanvas` or canvas at module top level. Workers, probes and wasm start in effects or on first call. CI imports every published entry in plain Node without a DOM.

## Public API

- **Async-only.** Every member that resolves arbitrary anchors or reads or writes the document engine returns a promise. Members that read main-thread state or a projection the worker pushes stay synchronous: zoom, the current page, the page count once laid out, search state, focus, subscriptions, the proposal geometry mirror, and the rendered-page geometry (`RenderedDomContext`) of the presented pages, read from the per-page line boxes with display spans that the worker pushes.
- **Async twins, then removal.** Each synchronous member outside that rule gets an async twin. A twin that returns a version is served by the authority that executes the edits consuming it, so a read and the edit built from it never straddle replicas. The synchronous member is marked `@deprecated` in the release that ships its twin and is removed at a flip at least one release after the last deprecation. Until then it works while the main replica is hydrated. On an unhydrated replica whose worker holds state the source lacks (worker-held proposals, and every edit once writes move to the worker), it throws a typed error naming its twin (`DocxReplicaNotReadyError`), since hydrating from the source would drop that state. A deprecated synchronous write on a hydrated replica reaches the worker as a host-origin update, in call order and with its history policy. A worker write resolves only after a hydrated replica has applied its update, so a deprecated synchronous read after an awaited write sees it; an unhydrated replica takes the current state when it hydrates.
- **Void commands.** Synchronous members that return nothing and only enqueue work (`highlightRange`, `resolveComment`) stay synchronous; the work runs in call order, and save awaits it.
- **Inventory.** DOCX: the `DocxEditorRef` members `getDocument`, `getPositionAtPoint`, `getSelectionInfo`, `getPageContent`, `findInDocument`, `scrollToParaId`, `scrollToCommentId`, `scrollToChangeId`, and the mutations that return results, `replyToComment` included; the document-valued `onContentChange` and `onChange` prop, whose twins hand listeners the version; plugin `getPositionAtPoint`, and `getAnchorGeometry` for targets the mirror does not carry; `getEditorRef`, whose `PagedEditorRef` exposes the `YrsSession`, is deprecated without a twin. PPTX and XLSX: `PptxEditorApi.handle`, `XlsxEditorApi.handle`, their synchronous `save(): Uint8Array`, and the plugins' `pluginAccess.handle`. Compile-time classification tables list every member of the editor ref, the editor props and the plugin context and geometry, so a new synchronous document member fails the typecheck.
- **No public inline runtime.** There is no `runtime: 'inline'` option. Until its flip, each format's worker path is an experimental opt-in (DOCX: `experimentalWorkerOpen`); at the flip the opt-in becomes a deprecated no-op.
- **Internal in-process transport.** jsdom and bun have neither `Worker` nor `OffscreenCanvas`. Tests run the same session host over an in-process transport that still structured-clones every message. It is not exported.

## Ownership

`shared/office-session/` is new and knows no format. It holds the protocol envelope, the client, the worker host, the surface registry, the capability probe, the in-process test transport and the import-closure test. Like `shared/host-contracts`, it is source-shared and bundled into each package. `packages/<fmt>/src/session/` holds each format's typed methods, its projections and a worker entry, a tsup entry like the one in `packages/docx/tsup.config.ts`. `packages/<fmt>-react` keeps UI, input, overlays and accessibility, fed by projections and queries. The Rust crates keep parse, edit, layout, display lists and save; this plan changes where they run.

## Reuse map

| Area | Reuse from DOCX | New work |
| --- | --- | --- |
| Client | `ResidentEngineWorkerClient`: pending map, terminal `fail`, silence watchdog (`RESIDENT_WORKER_SILENCE_MS`) | Typing by method table, events, `AbortSignal` for queries and exports |
| Worker host | The serial `operations` chain; a wasm trap is terminal (`residentEngineWorker.ts`) | Latest-wins lane for viewport updates |
| Surfaces | `attachCanvases` windowing, which zeroes off-window canvases because they cannot be transferred again; atomic back-buffer present | Per-format surface ids; thumbnail `ImageBitmap`s, closed when not consumed |
| Frames | FrameDelta epochs, damage and compact shifts (`packages/docx/src/layout/render/frameDelta.ts`) | PPTX and XLSX binary frames that never leave the worker |
| Text | Glyph-outline paint (`glyphCache.ts`, `crates/ooxml-text/src/outline.rs`) | An outline export for PPTX; glyph runs for XLSX |
| Images | The no-external-fetch rule (`canvasImageResolver.ts`) | Worker decode of package media with `createImageBitmap`, replacing main-thread `new Image()`; SVG through the `resvg` already in `crates/pptx-raster` |
| Open | The preview-then-full lifecycle and worker open | First slide for PPTX; for XLSX, an `open`/`finishCalculation` split |
| Reads | `documentRead` and the proposal registry the main thread mirrors | Per-format query tables |

## Shared runtime contract

```ts
interface OfficeSession<M extends SessionMethods, E extends SessionEvents> {
  open(input: { bytes: ArrayBuffer; fonts: FontInput[]; options: OpenOptions }): Promise<SessionState>;
  attachSurfaces(input: { surfaces: Array<{ id: SurfaceId; canvas: OffscreenCanvas }>; viewport: Viewport }): Promise<void>;
  updateViewport(viewport: Viewport): void;
  save(): Promise<ArrayBuffer>;
  dispose(): Promise<void>;
  readonly call: Promisified<M>;
  on<K extends keyof E>(event: K, listener: (payload: E[K]) => void): () => void;
}
```

- **Transport.** A minimal internal typed RPC over `postMessage` with transfer lists. The watchdog and terminal semantics already exist and are tested (`residentEngineWorkerClient.test.ts`); a mapped type over `M` gives client/host parity without a dependency.
- **Open.** `open` transfers the document and font bytes. It returns `SessionState { format, stage: 'preview' | 'ready', version, dirty, surfaces, capabilities }`. During `'preview'`, mutations and `save` are refused.
- **Surfaces.** Each canvas is transferred once, under a stable ordinal id. `updateViewport` carries `{ surface, x, y, width, height, zoom, dpr, overscan }`, is coalesced per animation frame, and the latest value wins. The worker builds and paints only the viewport plus overscan.
- **Ordering.** Calls go through one serial queue, with order-preserving input coalescing (batch sealing). Session and viewport generations let the worker skip presenting a frame for a superseded viewport; the client cannot un-present pixels. Geometry projections carry the epoch of the presented frame, and the client drops only obsolete projections; receipts and yrs updates are always delivered. Reorder tests cover this.
- **Save and dispose.** The session's `save` resolves after every queued mutation, reads the same authority that renders, and transfers an `ArrayBuffer`. The editor's `save` first runs its own input barrier (pending input, IME composition, gestures). `dispose` is idempotent: it rejects pending calls, releases surfaces and terminates the worker.
- **Failure.** Refusals are data (`OperationRefusal`, `shared/host-contracts/edits.ts`) and keep the session. A crash, a wasm trap, running out of memory, an unreadable message, or 60 s of silence while calls wait is terminal: pending calls reject with `SessionFailure`, the state becomes `failed`, and the editor shows it. There are no per-request timeouts.
- **Recovery.** The main thread retains the source bytes and a recovery log.
  - **yrs-backed sessions** (DOCX, PPTX, collaborative XLSX). Saved OOXML carries no CRDT clocks, so the log is a worker checkpoint (`encodeStateAsUpdate`) plus every update accepted since, local and remote, in order. Session state outside yrs is journaled in the same ordered log as the updates, so no acknowledged change is lost: DOCX comment replies and resolution, and every proposal registry change (creation with its id, retry key and suggestion metadata, decisions, withdrawal and `previewVersion`). The first checkpoint is taken at `ready`, before any edit is admitted. Each checkpoint names the last update it covers, and the log is truncated only after the worker acknowledges it. Reopen seeds a **fresh replica** with a new client id through the format's with-source entry (PPTX `open_from_update_with_source`; DOCX `openDocx` from the source, then `loadState`). A new id is required because PPTX `open_from_update` resets `id_counter`, and ids derive from the client id and that counter alone. The provider then reconnects with the new replica. Checkpoints restore through an internal path bounded by the open limits, not by the 64 MiB network-update cap, since a PPTX checkpoint carries the package and base64 media. Each checkpoint also carries that non-yrs state.
  - **Standalone XLSX sessions.** Collaborative mode refuses structural ops, so these stay standalone. The log is a journal of committed effects: the ops each commit applied, including those that undo, redo and accepted batches resolved to, with the commit's `now_serial` and the preserved-sheet state it adopted (`PreservedSheetState`). Replaying calls would not work: after reopen the undo stack is empty, and `applyEdits` refuses a recorded `expectVersion` against a fresh nonce. Reopen applies the journal to the source bytes the session opened, not to the last save, because a restored sheet references source parts a save may have dropped; the journal is kept across saves until an engine checkpoint of the model plus the preserved state exists.
  - **After reopen.** Acknowledged proposals come back with their ids, retry keys and decisions. The editor states what is lost: input and requests that were not acknowledged, and undo history. Editing then continues.
- **Fonts and images.** Fonts arrive with `open`. Every text primitive paints as glyph outlines, including the DOCX `fillText` path and PPTX placeholder labels; a missing outline paints the substitute face, never a blank, so there is no dependency on a worker `FontFaceSet`. Images decode through `createImageBitmap`. The closure test bundles every worker entry, including save, against a global without `document`, `window`, `Image` or `HTMLCanvasElement`; runtime tests decode, paint and export fixtures, SVG included, in that global.

## Per-format methods

| Format | Typed methods (`M`) | Projections back |
| --- | --- | --- |
| DOCX | `applyInput`, `applyDelete(direction, count)`, `setSelection`, `applyEdits`/`validateEdits`, `readParagraphs`, `findText`, `documentRead`, proposals, `hitTest(page, x, y)`, `rangeRects`, `verticalMove`, `selectionContext`, `exportStructuredWithPages`, comments, `undo`/`redo`, `applyUpdate` | Caret and selection rects, page sizes, the proposal mirror, receipts, damaged page ids |
| PPTX | `setActiveSlide`, `insertSlide`/`deleteSlide`/`moveSlide`, `insertText`/`deleteText`/`formatText`/`insertParagraphBreak`, `moveShape`/`resizeShape`/`addShape`/`addPicture`, comments, proposals, `applyEdits`, `hitTest(slide, x, y)`, `caretLines`, `thumbnails(ids, width)`, `exportPng`, `undo`/`redo`, `applyUpdate` | Slide list, selection geometry, thumbnails, receipts |
| XLSX | `setActiveSheet`, `editCell`, `editCells`, `applyOps`, `patchRangeStyle`, `setNumberFormat`, `moveChart`, `applyEdits`, `readCells(range)`, `cellAtPoint`, `rangeRects`, `calculationStatus`, `exportPng`, `undo`/`redo`, `applyUpdate` | Sheet info, content extent, merges near the viewport, visible cell text, `limitedCells`, receipts |

The names follow today's handles (`packages/pptx/src/wasm/loader.ts`, `packages/xlsx/src/wasm/loader.ts`, `packages/docx/src/yrs/residentEngineWorkerProtocol.ts`). There is no shared command union.

## DOCX

DOCX already runs the worker half of this contract behind `experimentalWorkerOpen`, with a main-thread replica for everything that still reads it. The slice order that removes the replica, the host engine and the fallback is in [docx-worker-document-owner.md](docx-worker-document-owner.md).

## PPTX

- **Scope.** Layout is slide-scoped (`SlideRenderer::layout_scoped_slide`), the wasm is DOM-free and the painter is canvas-only. The worker takes `openPresentation`, layout, `paintSlide` (typed to the 2D-context subset common to both canvas kinds), image decode, thumbnails, `slideToPng` and `save`. React keeps overlays, text input, gestures and accessibility; `EditorModel` shrinks to slide ids, sizes and selection, and the hit tests and display-list reads in `interactions.ts` become worker queries.
- **Layout and seeding.** `refreshAt` lays out every slide, including on each remote update; lay out the active slide first, then only the thumbnails in the visible strip. `seed_doc` stores the modeled package in yrs as JSON, inserted pictures carry base64 media in shared state, and `validate_doc` decodes it all again. Moving media to references is a collaboration schema change that needs versioning, media distribution, retention of the matching source, and staged-clone validation before adoption; it lands separately behind a schema version.
- **Updates, hit tests and text.** A remote update returns a full `DeckSnapshot`; add an internal variant that returns damaged slide ids, and keep the snapshot-returning signatures the native and Python bindings expose. `hitTest` reads the last laid-out slide, so every click relays out; keep one `RenderedSlide` per slide id. Text paints through `ctx.font`/`fillText`; paint outlines from the glyph and font ids the display list already carries.
- **API.** `CollaborationReplica` stays synchronous for the headless `PresentationHandle` and `WorkbookHandle`. A new `AsyncCollaborationReplica` fronts the session; providers accept either and stay on the main thread with their transport. `PptxEditorApi.handle`, its `save(): Uint8Array` and `pluginAccess.handle` get async twins over `api.session` and are deprecated, per [Public API](#public-api).

## XLSX

- **Into the worker:** `openWorkbook`, recalculation, `display_list`, charts, `paintDisplayList`, `save` and `exportPng`.
- **Viewport cells.** The main thread keeps the scroll spacer and sends `updateViewport` on each scroll frame. The worker paints the viewport plus overscan into clipped pane surfaces: the frozen corner, frozen rows, frozen columns and the scrolling body. Until the next present, the main thread translates only the scrolling regions. This replaces `doPaint`, which today makes synchronous wasm calls, reallocates the backing store and re-renders the editor on every scroll frame.
- **Geometry.** `autofit_rows` walks every cell twice per frame. Cache it per model epoch and invalidate it on cell, column-style and merge changes.
- **Recalculation.** Formatting, merges, `moveChart`, undo/redo and remote updates run `rebuild_and_recalc_all` and clone the model. Formatting should recalculate nothing, value edits should seed `recalc_after`, and formula edits should update graph edges first, as `edit_cells` does. Sheet, defined-name and table changes rebuild the graph; undo and remote updates classify their change sets the same way. A preview from the file's cached values needs an `open`/`finishCalculation` split, since `Session::open` recalculates before it returns. Cells without cached values render as pending.
- **Charts, React and accessibility.** `bump_model_epoch` clears the chart cache on every edit; invalidate a chart only when its read ranges, owner, source or theme change, and conservatively when a dependency is unknown. React keeps the cell editor, the formula bar, IME and selection (`packages/xlsx/src/selection/model.ts`). The accessible grid moves to `readCells` for the visible rows, since `buildA11yGrid` matches text by clip origin and misattributes spilled text.
- **Text.** XLSX text commands carry no glyph runs yet, so shaping and font registration are a prerequisite, with a visual-parity check.

## Design rules

1. **No `serde_json::Value` in Rust mid-paths.** Typed or binary boundaries. PPTX `seed_doc` and the per-frame XLSX display-list JSON are the remaining cases.
2. **Viewport-scoped, lazy display data and compact frame deltas.**
3. **One layout owner.** No duplicate layout, font relayout or double bootstrap.
4. **No idle whole-document work.** Hash the package once, with Web Crypto.
5. **Registries, not per-frame DOM queries.** XLSX `pointToCell` still calls `getBoundingClientRect` on every mousemove.
6. **Overlays windowed to the viewport, with worker-backed queries for off-window targets.** Tab, note links and range targets must keep working for off-window content.
7. **Stable ordinal ids, never derived from position,** for surfaces, pages, slides and grid tiles.
8. **Position shifts instead of suffix rewrites.**
9. **A preview first frame before full seeding.** Provisional state is marked partial; NUMPAGES stays blank until the full layout.
10. **A silence watchdog instead of hard timeouts.** No single feature may disable a fast path for the whole document.

## Perf harness and CI benchmark

- **Harness.** `e2e/perf/`: a Vite app, a Playwright driver for Chromium phase timings, and a Node bench for engine phases. A `?format=` parameter mounts any of the three editors. The shared client emits `session:open`, `session:first-present`, `session:ready` and `session:reply` marks. The `longtask` and `event` observers, CDP heap and RSS sampling, and traces stay.
- **Metrics, per format.** Bytes to first present and to ready; main-thread long tasks during open and per interaction; input-to-present p50/p95; worker reply time; frame and message bytes; heap; save time; DOM nodes. The #535 completion check is a trace with no `*_bg.wasm` samples on the renderer main thread after `open`.
- **Fixtures.** Size knobs on `scripts/create-demo-doc.ts`, `create-demo-deck.ts` and `create-demo-workbook.ts` generate 300 pages with notes, floats and tables; 200 slides with pictures and charts; and 100k cells with 10k formulas and 5 charts. Private documents stay in a local corpus.
- **CI.** `.github/workflows/perf.yml` runs on PRs that touch `packages/` or `crates/`. It interleaves A/B runs against the merge base on one runner, 5 runs each, and compares the minimums. It fails on a budget breach or a regression of more than 15%, and posts JSON and Markdown as an artifact and a PR comment.

## Phases

0. **Contract and harness.** Extract `shared/office-session` from the DOCX client and worker with no behavior change. Add the capability probe, the in-process test transport, the closure test, the Node import check, the harness with its budgets, and report-only CI. Every slice that changes behavior is measured with the harness and waits for it; additive API slices do not. The slice that turns on the construction error waits for the probe.
1. **DOCX.** The slices in [docx-worker-document-owner.md](docx-worker-document-owner.md), each behind `experimentalWorkerOpen`, ending in its default flip.
2. **PPTX.** The worker session behind an experimental opt-in, with the engine fixes, the async replica, the provider and the async API; then its flip.
3. **XLSX.** Glyph-run shaping, pane surfaces, the incremental recalculation classes, the `open`/`finishCalculation` split, charts, save and accessibility from `readCells`, behind an experimental opt-in; then its flip.

Each flip deletes that format's main-thread path and removes the synchronous members deprecated one release earlier. It is a breaking release for that editor, marked in its changesets. Headless handles do not change. VSDX is then evaluated on the same contract.

## Phase gate

Every phase passes `bun run test`, `bun run typecheck`, `bun run rust:check`, `bun run test:e2e` and `bun run test:e2e:browser`, the Node import check, and the Python bindings' own workspace checks (`cargo clippy` and `cargo test --manifest-path bindings/Cargo.toml --workspace`, plus the maturin-built Python e2e scenarios), since each binding's npm `typecheck` is `true`. Each slice adds the worker gates for what it implements, and a format's flip needs all of them:

- **Worker lifecycle.** Playwright specs cover transfer-once, window re-entry, dispose, stale-frame suppression, the construction error in a browser without `OffscreenCanvas`, and kill-the-worker, reopen and keep editing. After reopen, the new edits (typing, inserting a slide, editing a cell) converge with a peer and save. The spec also covers recovery across a save with undo, redo and an accepted batch; delete a sheet, save, undo, crash, reopen and save, with the sheet's preserved XML intact; a PPTX checkpoint over 64 MiB; and an unsaved DOCX comment reply and resolution, and an undecided proposal that keeps its retry identity.
- **Perf.** The format's budgets in `e2e/perf/budgets.json` hold, and no wasm runs on the main thread once the format's replica is gone.
- **Saved-output byte identity.** With clocks, client ids and seeds fixed (`RustSaveDeterminism` for DOCX, `now_serial` for XLSX), every fixture and scripted edit saves the same bytes through the worker as through the phase's base commit, except for separately approved baseline changes. Untouched parts and spans stay byte-identical, and `ooxml-diff` explains any difference.
- **Collaboration convergence.** `two-editors-converge` and `three-editors-mesh` pass, and a two-context browser proof with a worker editor and a headless peer reaches equal state vectors and equal saved bytes.

## Risks

- **Synchronous public APIs.** Hosts call them today. De-risk: async twins first, deprecation for one release, removal at the flip, and the member-listing test.
- **Unsupported browsers.** No fallback means an old browser cannot edit. De-risk: the construction error names the missing capability, the desktop target already supports it, and a WebKit project in the harness keeps it that way.
- **Crash data loss.** With one authority, a crash loses unsaved edits. De-risk: the recovery log and fresh-replica reopen, covered by the kill-the-worker spec.
- **Async collaboration.** Every provider encodes and applies updates synchronously and treats a return as success. De-risk: async state-vector and diff calls; applying incoming updates in order and awaiting validation; preserving origins; checking the connection generation after every await; and marking a peer synced only after its update applies.
- **Typing and IME latency.** De-risk: paint the caret in the worker, with an optimistic main-thread caret from the last geometry; coalesce input batches; keep IME composition in the DOM and commit once.
- **XLSX scroll latency.** De-risk: translate the scrolling pane surfaces over overscan, latest-viewport-wins, and a p95 scroll-to-present budget.
- **Memory.** Two heaps coexist until a format's replica goes, and the DOCX edit wasm reaches 2.5–3.1 GB. De-risk: remove the replica early, budget heap in the harness, and share one compiled `WebAssembly.Module`.
- **Accessibility.** Off-window content must stay reachable. De-risk: serve accessibility reads from worker queries rather than mounted nodes, with keyboard-reach specs in the gate.
- **Main-thread pixel consumers.** PPTX `presentFrame` hands plugins an `HTMLCanvasElement`, and XLSX print calls `window.print()` over the painted canvas. De-risk: plugins get worker `ImageBitmap`s and geometry, and print renders from `exportPng` bitmaps.
