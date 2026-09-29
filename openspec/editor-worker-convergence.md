# Editor worker convergence plan

This plan converges the DOCX, PPTX and XLSX web editors on one runtime, per [#535](https://github.com/openooxml/betteroffice/issues/535). A dedicated worker owns the only writable document handle from `open` to `dispose` and paints into transferred `OffscreenCanvas` surfaces. The main thread keeps React, focus, keyboard/IME, pointer input, overlays and the accessible projection. The formats share the runtime boundary, not the document model. Headless, native and Python APIs are unaffected.

## Target architecture

- **One authority.** The worker holds the wasm handles, undo, layout and recalculation caches, fonts, decoded images and save. The main thread holds no engine, no mirror and no display list. It receives small projections: state, version, dirty flag, surface sizes, caret/selection/hit geometry, query results, change events and yrs updates.
- **One render path.** Layout and paint share the worker, so display lists never cross the boundary. Thumbnails and exports come back as transferred `ImageBitmap`s or bytes.
- **One failure rule.** A worker failure ends the session visibly, and nothing takes over silently. The compatibility path the maintainer asked for on #535 is an explicitly selected `runtime: 'inline'` adapter. It runs the same session host in-process, is chosen at construction, and never takes over mid-session.

## Today

| Stage | DOCX | PPTX | XLSX |
| --- | --- | --- | --- |
| Open | Main thread: `seedYrsSession` → `openDocx` (`packages/docx-react/src/components/DocxEditor/hooks/useYrsCoreSession.ts:119`) | Main thread: `openPresentation` (`packages/pptx/src/wasm/loader.ts:332`) | Main thread: `openWorkbook` (`packages/xlsx/src/wasm/loader.ts:551`) → `Workbook::open_recalculated` (`crates/betteroffice-xlsx/src/workbook.rs:392`) |
| Layout / recalc | Main-thread `computeLayout` (`hooks/useLayoutPipeline.ts:339`), then replayed by the worker's `hydrate` (`packages/docx/src/yrs/residentEngineWorker.ts:232`) | Main thread, for every slide, on open and on each remote update (`refreshAt`, `packages/pptx-react/src/PptxEditor.tsx:582-651`) | Main thread, inside each mutating call; ops, undo and remote updates run `rebuild_and_recalc_all` (`crates/xlsx-calc/src/engine.rs:52`) |
| Display list | Worker, as a binary FrameDelta | JSON per slide (`crates/pptx-wasm/src/lib.rs:169`) | JSON per viewport, rebuilt every rAF (`crates/xlsx-wasm/src/core.rs:325`) |
| Image decode | Main-thread `new Image()` (`packages/docx/src/layout/render/canvasImageResolver.ts:24`); pages with images leave the worker path (`CanvasPagesView.tsx:232`) | Main-thread `createImageBitmap`; SVG via `new Image()` (`packages/pptx/src/render/image.ts:31-49`) | None; pictures are not painted |
| Paint | Worker `OffscreenCanvas`, back-buffered (`residentEngineWorker.ts:365-465`) | Main thread, straight into the visible canvas (`PptxEditor.tsx:1041-1068`; `paintSlide`, `render/canvas.ts:67`) | Main thread, the whole frame on every scroll (`doPaint`, `packages/xlsx-react/src/XlsxEditor.tsx:1186`; `paintDisplayList`, `packages/xlsx/src/render/canvas2d.ts:39`) |
| Save | Main-thread `saveYrsDocx` (`packages/docx/src/yrs/saveYrsDocx.ts:231`) | Main thread, synchronous `handle.save()` (`loader.ts:553`) | Main thread, synchronous `handle.save()` (`loader.ts:841`) |
| React holds | Host `YrsSession`, the display-list snapshot, the worker client and the fallback engine (`hooks/useDisplayList.ts:159-200`) | The handle, plus `EditorModel` with the deck snapshot, the frame and every thumbnail list (`PptxEditor.tsx:236-243`) | The handle, the full `frame`, `sheetInfo`, selection and `revision` (`XlsxEditor.tsx:601-695`) |
| Off-thread | The resident engine worker, the repository's only `new Worker` (`packages/docx/src/yrs/residentEngineWorkerClient.ts:68`) | None; `OffscreenCanvas` appears only as main-thread scratch (`render/canvas.ts:244`, `render/png.ts:35`) | None; `packages/xlsx/src/headless.ts` is DOM-free math with no wasm, and its purity test checks only import specifiers (`src/__tests__/seam-purity.test.ts:10`) |

## Ownership

`shared/office-session/` is new and knows no format. It holds the protocol envelope, the client, the worker host, the surface registry, the inline adapter and the import-closure test. Like `shared/host-contracts`, it is source-shared and bundled into each package. `packages/<fmt>/src/session/` holds each format's typed methods, its projections and a worker entry, which is a tsup entry like the one in `packages/docx/tsup.config.ts:37`. `packages/<fmt>-react` keeps UI, input, overlays and accessibility, fed by projections and queries. The Rust crates keep parse, edit, layout, display lists and save; this plan changes where they run and adds the engine fixes the design rules call for.

## Reuse map

| Area | Reuse from DOCX | New work |
| --- | --- | --- |
| Client | `ResidentEngineWorkerClient`: pending map, terminal `fail`, #896 silence watchdog (`RESIDENT_WORKER_SILENCE_MS`) | Typing by method table, events, `AbortSignal` for queries and exports |
| Worker host | The serial `operations` chain; a wasm trap is terminal (`residentEngineWorker.ts:59-82`) | Latest-wins lane for viewport updates |
| Surfaces | `attachCanvases` windowing, which zeroes off-window canvases because they cannot be transferred again (`:157-189`); atomic back-buffer present (`:365-465`) | Per-format surface ids; thumbnail `ImageBitmap`s, closed when not consumed |
| Frames | FrameDelta epochs and damage (`packages/docx/src/layout/render/frameDelta.ts`); compact shifts (#918) | PPTX and XLSX binary frames that never leave the worker |
| Text | Glyph-outline paint (`packages/docx/src/layout/render/glyphCache.ts`, `crates/ooxml-text/src/outline.rs:138`) | An outline export for PPTX; glyph runs for XLSX |
| Images | The worker resolver from closed #729 (`fetch` + `createImageBitmap`) | SVG, which today needs `new Image()` |
| Open | The #909/#914 preview-then-full lifecycle | First slide for PPTX; cached values before recalculation for XLSX |

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

- **Transport.** A minimal internal typed RPC, not kkrpc. The watchdog, transfer lists and terminal semantics already exist and are tested (`residentEngineWorkerClient.test.ts`), and a mapped type over `M` gives client/host parity without adding a dependency.
- **Open.** `open` transfers the document and font bytes. It returns `SessionState { format, stage: 'preview' | 'ready', version, dirty, surfaces, capabilities }`. During `'preview'`, mutations and `save` are refused, as #914 does.
- **Surfaces.** Each canvas is transferred once, under a stable ordinal id. `updateViewport` carries `{ surface, x, y, width, height, zoom, dpr, overscan }`, is coalesced per animation frame, and the latest value wins. The worker builds and paints only the viewport plus overscan.
- **Ordering.** Calls go through one serial queue, with order-preserving input coalescing (#896's batch sealing). Replies carry `version` and `frameEpoch`, and the client drops anything older than what it holds; reorder tests cover it, since #901 once painted page 8's text onto page 7. Mutations are never cancelled.
- **Save and dispose.** `save` resolves after every queued mutation, reads the same authority that renders, and transfers an `ArrayBuffer`. `dispose` is idempotent: it rejects pending calls, releases surfaces and terminates the worker.
- **Failure.** Refusals are data (`OperationRefusal`, `shared/host-contracts/edits.ts`) and keep the session. A crash, a wasm trap, an unreadable message, or 60 s of silence while calls wait is terminal: pending calls reject with `SessionFailure`, the state becomes `failed`, and the editor shows it. There are no per-request timeouts.
- **Recovery.** Recovery is explicit, never a hidden second engine. The main thread keeps the local yrs updates the worker has emitted since the last save; collaboration already consumes this stream. "Reopen" is `open(lastSavedBytes)` followed by `applyUpdate(log)`.
- **Fonts and images.** Fonts arrive with `open` and paint as glyph outlines, so there is no dependency on a worker `FontFaceSet`. Images decode through `createImageBitmap`. The closure test bundles every worker entry, including the save modules, against a global without `document`, `window`, `Image` or `HTMLCanvasElement`.

## Per-format methods

| Format | Typed methods (`M`) | Projections back |
| --- | --- | --- |
| DOCX | `applyInput`, `applyDelete(direction, count)`, `setSelection`, `applyEdits`/`validateEdits`, `readParagraphs`, `findText`, `hitTest(page, x, y)`, `rangeRects`, `verticalMove`, `selectionContext`, `exportStructuredWithPages`, `undo`/`redo`, `applyUpdate` | Caret and selection rects, page sizes, receipts, damaged page ids |
| PPTX | `setActiveSlide`, `insertSlide`/`deleteSlide`/`moveSlide`, `insertText`/`deleteText`/`formatText`/`insertParagraphBreak`, `moveShape`/`resizeShape`/`addShape`/`addPicture`, comments, proposals, `applyEdits`, `hitTest(slide, x, y)`, `caretLines`, `thumbnails(ids, width)`, `exportPng`, `undo`/`redo`, `applyUpdate` | Slide list, selection geometry, thumbnails, receipts |
| XLSX | `setActiveSheet`, `editCell`, `editCells`, `applyOps`, `patchRangeStyle`, `setNumberFormat`, `moveChart`, `applyEdits`, `readCells(range)`, `cellAtPoint`, `rangeRects`, `calculationStatus`, `exportPng`, `undo`/`redo`, `applyUpdate` | Sheet info, content extent, merges near the viewport, visible cell text, `limitedCells`, receipts |

The names follow today's handles: `packages/pptx/src/wasm/loader.ts:556-792`, `packages/xlsx/src/wasm/loader.ts:358-487` and `packages/docx/src/yrs/residentEngineWorkerProtocol.ts`. There is no shared command union.

## DOCX convergence gap

- **The host is still the authority.** `seedYrsSession` opens the document on the main thread. The worker is a replica bootstrapped from `residentWorkerSnapshot` (`packages/docx/src/yrs/index.ts:1524`), and its edits come back as yrs updates that the host applies. The host also runs `computeLayout`, the font relayout (`useLayoutTriggers.ts:54`), queries (`displayListQueries.ts:423`) and save (`saveYrsDocx.ts:231`, `useFileIO.ts:102`).
- **The main-thread fallback still exists.** `fallback` (`useDisplayList.ts:730`) and `replayInputOnMainThread` (`:399`) destroy the worker and pin `workerFallbackEngineRef` for the rest of the session. Pages with images also bypass the worker, because `displayListNeedsHostImages` sends them to the DOM canvas.
- **#896 changed the fallback semantics.** Before it, the 5 s and 15 s per-request timeouts (`REQUEST_TIMEOUT_MS`, `residentEngineWorkerClient.ts:50`) fired on slow keys in large documents, and editing moved to the main thread for good. Since #896, only a crash, a trap, an unreadable message, or 60 s of silence (reset by every reply) fails the worker, and the recovery frame is full with continuing epochs (`resetFrameBase`). The fallback is therefore crash recovery only. Replacing it with terminal failure plus a reopen from the update log costs no speed; the only loss is surviving a crash silently.
- **Path.** First, #900 (the first layout owned by the worker) and #909/#914 (the preview) become `open` in `'preview'`, through an `openPreview`/`finishOpen` split that transfers bytes instead of a snapshot. Next, queries (#899, #913), images and save move into the worker. Last, the host engine, `residentWorkerSnapshot`, `fallback` and `replayInputOnMainThread` are deleted.

## PPTX migration (first slice)

- **Why PPTX first.** It is the smallest full slice: layout is slide-scoped (`SlideRenderer::layout_scoped_slide`, `crates/pptx-render/src/layout.rs:381`), the wasm is DOM-free, and the painter is canvas-only. The worker takes `openPresentation`, layout, `paintSlide` (typed to the 2D-context subset common to both canvas kinds), image decode, thumbnails, `slideToPng` and `save`.
- **What stays in React.** Overlays, text input, gestures and accessibility. `EditorModel` shrinks to slide ids, sizes and selection. The hit tests and display-list reads in `interactions.ts` become worker queries.
- **Layout and seeding fixes.** `refreshAt` lays out every slide (`PptxEditor.tsx:599-605`), including on each remote update (`:919-921`); instead, lay out the active slide first, then only the thumbnails in the visible strip. `seed_doc` stores the whole package in yrs as JSON with base64 media (`crates/pptx-edit/src/deck.rs:44`), and `validate_doc` decodes it again (`:959`); instead, keep media by reference and the model typed.
- **Update, hit-test and text fixes.** A remote update returns a full `DeckSnapshot` (`crates/pptx-edit/src/lib.rs:327-351`); return the damaged slide ids instead. `hitTest` reads the last laid-out slide, so every click relays out (`PptxEditor.tsx:1446`); keep one `RenderedSlide` per slide id. Text paints through `ctx.font`/`fillText` (`render/canvas.ts:936-960`); paint outlines from the glyph and font ids that the display list already carries.
- **API.** `PptxEditorApi.handle`, its synchronous `save(): Uint8Array` (`PptxEditor.tsx:154-189`) and `CollaborationReplica` (`packages/pptx/src/collaboration/types.ts:3`) become async. The provider and transport stay on the main thread and talk to an async replica.

## XLSX migration

- **Into the worker:** `openWorkbook`, recalculation, `display_list` (`workbook.rs:1886`), charts (`crates/xlsx-render/src/chart.rs:381`), `paintDisplayList`, `save` and `exportPng`.
- **Viewport cells.** The main thread keeps the scroll spacer and sends `updateViewport` on each scroll frame. The worker paints the viewport plus overscan into one grid surface, and the main thread translates that surface until the next present. This replaces `doPaint`, which today makes synchronous wasm calls, reallocates the backing store and re-renders the editor on every scroll frame.
- **Geometry.** `autofit_rows` walks every cell twice per frame (`crates/xlsx-render/src/geometry.rs:167`, via `workbook.rs:3895`). Cache it per model epoch and update only the rows an edit touches.
- **Recalculation.** Today `apply_ops` (formatting, merges and `moveChart`; `workbook.rs:1388`), undo/redo (`:1502`) and remote updates (`:788`) run `rebuild_and_recalc_all` and clone the model (`:1306`, `:1401`, `:576`). Formatting should recalculate nothing, and the rest should seed `recalc_after` (`engine.rs:39`). The first frame paints the file's cached values as `'preview'` before the open-time `recalculate_all`.
- **Charts, React and accessibility.** `bump_model_epoch` (`workbook.rs:2861`) clears the chart cache on every edit; invalidate each chart by the ranges it reads instead. React keeps the cell editor, the formula bar, IME and selection, which is already pure TS (`packages/xlsx/src/selection/model.ts`). The accessible grid moves to `readCells` for the visible rows: `buildA11yGrid` matches text by clip origin (`packages/xlsx/src/a11y/index.ts:56-62`), which needs the whole command stream and misattributes spilled text.

## Design rules

These rules come from the DOCX large-document work. #911 integrates #894–#925; on a 267-page document it takes the first canvas from 8.6–8.9 s to 1.64–1.69 s, or 0.61–0.67 s with the preview.

1. **No `serde_json::Value` in Rust mid-paths.** In #894 the fingerprint went from 0.81 s to under 50 ms; #897 does the same for another path. Next targets: PPTX `seed_doc` and the XLSX JSON display list.
2. **Viewport-scoped, lazy display data and compact frame deltas.** In #901 the first frame went from 131–151 MB to 0.9 MB; in #899 the heap went from 1071 MB to 636 MB. See also #898, #913, #919 and #920.
3. **One layout owner.** #900 removed the duplicate layout, the font relayout and the double bootstrap.
4. **No idle whole-document work.** #905 removed a 1.6 s idle `materializeDocx`, and #908 a 3,799-story pass per insert; #902 and #921 hash the package once, with Web Crypto.
5. **Registries, not per-frame DOM queries.** Before #910, `querySelector` cost 1.44 s per caret update. XLSX `pointToCell` still calls `getBoundingClientRect` on every mousemove (`XlsxEditor.tsx:1480`).
6. **Overlays windowed to the viewport, with worker-backed queries for off-window targets.** #915 cut table handles from 575 to at most 21. #917, which cut DOM nodes from 634k to 6.7k, is held back because Tab, note links and `findElementsForRange` broke for off-window content.
7. **Stable ordinal ids, never derived from position (#916).** This applies to surfaces, pages, slides and grid tiles.
8. **Position shifts instead of suffix rewrites.** In #918 an early keystroke's frame went from 5.5 MB to 283 KB; see also #925 and #912.
9. **A preview first frame before full seeding (#909, #914).** Provisional state is marked partial; with #907, NUMPAGES stays blank until the full layout.
10. **A silence watchdog instead of hard timeouts (#896).** In the same spirit, no single feature may disable a fast path for the whole document (#903, #906, #923).

## Perf harness and CI benchmark

- **Harness.** The harness behind #894–#925 moves into `e2e/perf/`: a Vite app, a Playwright driver for Chromium phase timings, and a Node bench for engine phases. A `?format=` parameter mounts any of the three editors. The absolute paths and the `Worker` monkeypatch go; the shared client emits `session:open`, `session:first-present`, `session:ready` and `session:reply` marks instead. The `longtask` and `event` observers, CDP heap and RSS sampling, and traces stay.
- **Metrics, per format.** Bytes to first present and to ready; main-thread long tasks during open and per interaction; input-to-present p50/p95; worker reply time; frame and message bytes; heap; save time; and DOM nodes. The #535 completion check is a trace that shows no `*_bg.wasm` samples on the renderer main thread after `open`, under `runtime: 'worker'`.
- **Fixtures.** Size knobs on `scripts/create-demo-doc.ts`, `create-demo-deck.ts` and `create-demo-workbook.ts` generate three fixtures: 300 pages with notes, floats and tables; 200 slides with pictures and charts; and 100k cells with 10k formulas and 5 charts. The real 267-page document stays in a local corpus.
- **CI.** `.github/workflows/perf.yml` runs on PRs that touch `packages/` or `crates/`. It interleaves A/B runs against the merge base on one runner, 5 runs each, and compares the minimums. It fails on a budget breach or on a regression of more than 15%, and posts JSON and Markdown as an artifact and a PR comment. DOCX budgets start from #911's numbers; PPTX and XLSX budgets start from the phase 0 baseline.

## Phases

0. **Contract and harness.** Extract `shared/office-session` from the DOCX client and worker with no behavior change. Add the closure test, the inline adapter, the harness and report-only CI.
1. **PPTX.** Ship the worker session behind `runtime: 'worker'`, together with the engine fixes and the async replica and API. Flip the default once the gate passes.
2. **DOCX.** Move open, images, queries and save into the worker. Remove the host engine and the automatic fallback; `'inline'` becomes the compatibility path.
3. **XLSX.** Ship the viewport surface, incremental recalculation, charts, save, and accessibility from `readCells`. Flip the default.
4. **Cleanup.** Drop the main-thread engine paths from the default bundles, and evaluate VSDX on the same contract.

## Phase gate

A phase is done when five checks pass. **Tests:** `bun run test`, `bun run typecheck` and `bun run rust:check`. **E2E:** `bun run test:e2e` and `bun run test:e2e:browser`, plus new Playwright specs for transfer-once, window re-entry, dispose, stale-reply drop and kill-the-worker-then-reopen. **Perf:** the format's budgets in `e2e/perf/budgets.json` hold, and no wasm runs on the main thread. **Saved-output byte identity:** every fixture and scripted edit saves identical bytes through the worker, through the inline adapter and at the phase's base commit, with `ooxml-diff` explaining any difference. **Collaboration convergence:** `two-editors-converge` and `three-editors-mesh` pass, and a two-context browser proof with one worker editor and one inline editor reaches equal state vectors and equal saved bytes.

## Risks

- **Synchronous public APIs.** These include `DocxEditorRef.getDocument`, `getPositionAtPoint` and `getSelectionInfo`; `PptxEditorApi.handle` and `XlsxEditorApi.handle`, with their `save(): Uint8Array`; and the plugins' `pluginAccess.handle`. De-risk: add async twins and serve cheap reads from the last projection; keep the synchronous handle only under `'inline'`, for one minor release, with changesets; and add a test that lists every synchronous member.
- **Browser baseline.** The desktop app targets `safari16.4` (`apps/desktop/vite.config.ts:5`). De-risk: probe capabilities once at construction, paint glyph outlines so worker font loading is never needed, and add a WebKit project to the harness.
- **Crash data loss.** With a single authority, a crash loses unsaved edits. De-risk: the update log plus reopen, covered by a kill-worker end-to-end test.
- **Typing and IME latency.** De-risk: paint the caret in the worker, as DOCX's painted-caret machine does, with an optimistic main-thread caret from the last geometry; coalesce input batches (#896); and keep IME composition in the DOM, committing once.
- **XLSX scroll latency.** De-risk: translate an overscan back buffer, apply latest-viewport-wins, and set a p95 scroll-to-present budget.
- **Memory.** Two heaps coexist during migration, and the DOCX edit wasm already reaches 2.5–3.1 GB. De-risk: remove the host engine early, budget heap in the harness, and share one compiled `WebAssembly.Module`.
- **Accessibility.** #917 broke Tab, links and range targets for off-window content. De-risk: serve accessibility reads from worker queries rather than mounted nodes, and add keyboard-reach specs to the gate.
- **Main-thread pixel consumers.** PPTX `presentFrame` hands plugins an `HTMLCanvasElement`, and XLSX print calls `window.print()` over the painted canvas. De-risk: plugins get worker `ImageBitmap`s and geometry, and print renders from `exportPng` bitmaps.
