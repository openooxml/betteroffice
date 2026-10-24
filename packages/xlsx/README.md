# @betteroffice/xlsx

Framework-free XLSX parsing, formula calculation, editing, and rendering through a Rust/WebAssembly
engine. Includes canvas painting, viewport, hit-testing, accessibility, and clipboard helpers.

## Install

```bash
bun add @betteroffice/xlsx
```

No peer dependencies.

## Usage

Sheet, row and column indices are zero-based, so `editCell(0, 9, 2, ...)` writes cell C10 on the
first sheet.

```ts
import { initWasm, openWorkbook } from '@betteroffice/xlsx';

export async function updateWorkbook(bytes: Uint8Array): Promise<Uint8Array> {
  await initWasm();
  const workbook = openWorkbook(bytes);
  try {
    workbook.editCell(0, 9, 2, '=SUM(C1:C9)');
    return workbook.save();
  } finally {
    workbook.dispose();
  }
}
```

Cell input parses text, numbers, booleans, and formulas;
formulas start with `=`. Editing a cell recalculates its dependents.

## Entry points

These entry points are exported from `@betteroffice/xlsx`:

| entry point | description |
| --- | --- |
| `initWasm(input?)` | Initialize the engine; concurrent calls share initialization. |
| `openWorkbook(bytes, options?)` | Open XLSX bytes as a synchronous `WorkbookHandle` after initialization. |
| `openWorkbookSession(bytes, options?)` | Open a workbook in a dedicated worker with asynchronous access. |
| `paintDisplayList(ctx, frame, dpr, origin?)` | Paint a display list onto a canvas 2D context. |
| `isPngExportAvailable()` | Check whether the initialized engine supports PNG export. |
| `isProposalsAvailable()` | Check whether the initialized engine supports edit proposals. |
| `exportXlsxStructured(bytes, options?)` | Asynchronously export stored workbook content from bytes. |
| `exportXlsxMarkdown(bytes, options?, markdownOptions?)` | Asynchronously export Markdown from bytes. |

`initWasm()` loads the packaged WASM asset by default. Pass WASM bytes or a compiled
`WebAssembly.Module` in runtimes that cannot fetch that URL. Dispose workbook handles when finished.

## Canvas rendering and grid helpers

`WorkbookHandle.displayList(viewport)` returns the active sheet's display list. Size the backing store
for the device pixel ratio and give the canvas its logical CSS dimensions:

```ts
import { paintDisplayList, type WorkbookHandle } from '@betteroffice/xlsx';

export function drawWorkbook(workbook: WorkbookHandle, canvas: HTMLCanvasElement) {
  const dpr = window.devicePixelRatio || 1;
  canvas.width = 800 * dpr;
  canvas.height = 600 * dpr;
  canvas.style.width = '800px';
  canvas.style.height = '600px';
  const ctx = canvas.getContext('2d');
  if (!ctx) throw new Error('Canvas 2D is unavailable');
  const frame = workbook.displayList({ x: 0, y: 0, width: 800, height: 600 });
  paintDisplayList(ctx, frame, dpr);
}
```

For a custom grid, use `cellAtPoint`, `cellRect`, and `rangeRect` for hit-testing and geometry,
`buildA11yGrid` for accessibility data, and `toTsv` / `fromTsv` for clipboard text.
The `@betteroffice/xlsx/headless` entry exports pure grid helpers and types without the canvas painter
or WASM loader.

## Range moves

`workbook.moveRange(sheet, source, destination)` moves an A1 range within a sheet and rewrites
references in one undo step; the destination is the new top-left cell. Surface errors when a move is
refused rather than discarding the source workbook.

## Printing and PNG export

`workbook.printDisplayList(sheet, range, metrics, gridlines)` renders an A1 range using `PrintMetrics`
measured from the workbook's normal font. The caller chooses paper size, margins, and pagination;
`paintDisplayList` accepts an optional origin in backing-store pixels for positioning the result.

`workbook.renderPng(viewport)` exports the active sheet's viewport. `workbook.renderRangePng()` exports
its used range, or accepts `{ range, scale }` to select an A1 range and raster scale. Both return PNG
bytes synchronously; check `isPngExportAvailable()` before calling them.

## Worker sessions

The low-level `openWorkbookSession` API (also exported from `@betteroffice/xlsx/worker`) is
experimental. It returns a `WorkbookSession` with
promise-returning operations through `session.call`, plus asynchronous `session.save()` and
`session.dispose()` methods.

The React editor uses a worker by default and exposes its own asynchronous API. Use
[`@betteroffice/xlsx-react`](https://www.npmjs.com/package/@betteroffice/xlsx-react) for the editor UI.

## Edit proposals

`workbook.propose(agentId, note, edits)` stages cell edits for review without applying them.
`listProposals`, `acceptProposal`, and `rejectProposal` manage that queue; acceptance applies the
proposal as one undo step and recalculates dependents. Check `isProposalsAvailable()` first;
acceptance throws `StaleProposalError` if targeted cells have changed since staging.

## Version-checked edit batches

`readCells` returns cells with their version; `validateEdits` checks a batch without committing it,
and `applyEdits` applies all steps or returns a typed refusal without changing the workbook.
See [Edit XLSX in version-checked batches](https://docs.betteroffice.dev/docs/javascript#edit-xlsx-in-version-checked-batches)
for targets, guards, supported steps, and limits.

## Structured export

`exportStructured` and `exportMarkdown` read committed content with positional anchors and diagnostics
without recalculating formulas. See
[Export XLSX content with anchors](https://docs.betteroffice.dev/docs/javascript#export-xlsx-content-with-anchors)
for scopes, hidden content, limits, and byte-based exports.

## Collaboration

The `@betteroffice/xlsx/collaboration` entry exports `CollaborationProvider` for connecting replicas
to a binary transport. See the [collaboration guide](https://docs.betteroffice.dev/docs/collaboration)
for provider lifecycle, transport setup, and shared starting state.

Open the workbook with `openWorkbook(bytes, { collaborative: true })`. Collaborative workbooks
synchronize cell content, formulas, styles, column widths and row heights, with local undo and
redo whose effects synchronize.
Structural edits throw: inserting or deleting rows and columns, adding, removing, moving or renaming
sheets, and changing charts or defined names.

## Operation profiles

`editCellProfiled`, `applyOpsProfiled`, and `displayListProfiled` return engine stage durations in
milliseconds alongside their results. See
[Operation profiles](https://docs.betteroffice.dev/docs/javascript#operation-profiles).

## Documentation

- [JavaScript guide](https://docs.betteroffice.dev/docs/javascript)
- [Package guide](https://docs.betteroffice.dev/docs/packages)
- [React editor on npm](https://www.npmjs.com/package/@betteroffice/xlsx-react)
- [Repository](https://github.com/openooxml/betteroffice)
