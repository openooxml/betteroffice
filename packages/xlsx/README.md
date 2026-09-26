# @betteroffice/xlsx

Framework-free core for the BetterOffice XLSX editor — the Rust engine (parse,
calc, render) compiled to WebAssembly, plus display-list, viewport, hit-test,
and accessibility helpers.

```bash
bun add @betteroffice/xlsx
```

Most apps want the turnkey React component in
[`@betteroffice/xlsx-react`](https://www.npmjs.com/package/@betteroffice/xlsx-react).
Reach for this package directly to render a spreadsheet onto your own canvas or
to drive a workbook headlessly.

## Open, render, edit, save

```ts
import { initWasm, openWorkbook, paintDisplayList } from "@betteroffice/xlsx";

await initWasm();
const workbook = openWorkbook(new Uint8Array(await file.arrayBuffer()));

const ctx = canvas.getContext("2d")!;
const dpr = devicePixelRatio;
canvas.width = 800 * dpr;
canvas.height = 600 * dpr;
const frame = workbook.displayList({ x: 0, y: 0, width: 800, height: 600 });
paintDisplayList(ctx, frame, dpr);

workbook.editCell(0, 9, 2, "=SUM(C1:C9)"); // recalcs dependents
const bytes = workbook.save();
```

`initWasm()` fetches the packaged wasm asset once in browsers. Pass wasm bytes
or a precompiled `WebAssembly.Module` explicitly in runtimes that cannot fetch
the asset URL.

Around the handle, the package exports the helpers a custom grid needs:
`cellAtPoint` / `cellRect` / `rangeRect` (hit-testing), the viewport math,
`buildA11yGrid` (accessibility tree), and `toTsv` / `fromTsv` (clipboard). The
handle itself covers styling (`patchRangeStyle`, `setNumberFormat`), undo/redo,
and PNG export (`renderPng` / `renderRangePng`; guard with
`isPngExportAvailable`).

## Print a range

`workbook.printDisplayList(sheet, range, metrics, gridlines)` renders a range
without changing workbook data, the active sheet, or screen geometry. Pass
`PrintMetrics` measured from the workbook's normal font: `dpi`, `maxDigitWidth`,
`fontAscent`, and `fontDescent` use the layout device's pixels; `fontSizePt` and
`defaultRowHeightPt` use points. `fontFamily` supplies the default face, and
optional `defaultColumnWidth` uses the stored OOXML character width.

The result uses 96-DPI logical coordinates. Paint it with `paintDisplayList(ctx,
frame, scale, { x, y })`; the optional origin uses backing-store pixels, so page
margins need no intermediate image. Cell ranges, paper size, and pagination are
chosen by the caller. See the [Office comparison harness](../../scripts/office-quality).

## AI agents / human-in-the-loop

An agent stages edits as a proposal instead of applying them; a human reviews
per-cell before/after previews and accepts or rejects. Accepting applies the
proposal as one undo step and recalcs dependents, and throws
`StaleProposalError` if the workbook drifted under the proposal since it was
staged.

```ts
import { isProposalsAvailable } from "@betteroffice/xlsx";

if (isProposalsAvailable()) {
  const proposal = workbook.propose("copilot", "add totals", [
    { sheet: 0, row: 9, col: 2, input: "=SUM(C1:C9)" },
  ]);
  workbook.listProposals(); // pending proposals, oldest first
  workbook.acceptProposal(proposal.id); // or workbook.rejectProposal(proposal.id)
}
```

The React editor paints pending proposals as in-cell tracked-change ghosts with
an accept/reject panel. Guard with `isProposalsAvailable()` against cores built
without the feature. `StaleProposalError.targets` names each drifted cell's sheet
beside `cells`.

## Version-checked edit batches

Read cells with the version they were read at, then apply a batch against it.
Every step commits as one recalculated change and one undo step, or the batch
returns a typed refusal and nothing changes:

```ts
const read = workbook.readCells({
  ranges: [{ sheetId: "sheet:0", range: { kind: "a1", a1: "B3" } }],
});
if (!read.ok) throw new Error(read.failure.message);

const result = workbook.applyEdits({
  expectVersion: read.version,
  steps: [
    {
      op: "setCellInputs",
      target: { sheetId: "sheet:0", range: { kind: "a1", a1: "B3" } },
      inputs: [["120"]],
      expect: { cells: [[{ value: read.ranges[0].cells[0][0].value }]] },
    },
    { op: "setNumberFormat", target: { sheetId: "sheet:0", range: { kind: "a1", a1: "B3" } }, format: "currency" },
  ],
});
if (!result.ok) console.warn(result.failure.code); // e.g. "stale-version"
```

- Steps: `setCellInputs` (parsed like typing, against each cell's current
  number format), `setFormulas` (source without `=`, stored as formulas whatever
  the format), `setNumberFormat` and `patchStyle`. Content and formatting may
  combine on the same cells; writing one property twice refuses with
  `overlapping-steps`.
- Targets name a sheet id from the current catalog (`sheet:{index}` standalone,
  the replica's sheet keys in collaboration) and an A1 range or zero-based
  corners. Matrices and guards match the target's shape exactly. Guards compare
  a cell's value, formula (`null` for none) or display text before the batch.
- `validateEdits` stages a batch without changing anything; `findText` searches
  display text exactly and case-sensitively. Update listeners run once, after
  `applyEdits` returns, and see the recalculated state and its new version.
  `changedSheets` names every sheet the batch or its recalculation changed.
- Requests over 16 MiB and results over 64 MiB refuse with `limit-exceeded`;
  calculation diagnostics stop at 10,000 cells per list and set `truncated`.
- `history: "none"` keeps a batch out of undo; standalone undo still replays
  older steps over its cells. It is experimental: that interaction may change
  in a minor release. `source` records provenance only. Volatile functions see
  only `calculation.nowSerial`.
- Versions and sheet ids are session-scoped; standalone sheet ids are
  positional, so each is valid only for the version it was read at. Batches do
  not insert or delete rows, columns or sheets, merge cells or move charts, and
  refuse writes to merged-cell followers, array-formula cells and protected
  sheets.

## Structured export

XLSX exports bounded sparse worksheet content and Markdown with positional
anchors, formulas, stored values, formatted text, explicit hidden-content
options, and omission diagnostics. Export does not recalculate formulas:

```ts
const result = workbook.exportStructured({ scope: [{ sheet: 0, range: "A1:D20" }] });
if (result.ok) {
  for (const cell of result.content.sheets[0].cells) {
    console.log(cell.anchor, cell.value, cell.formula, cell.displayText);
  }
}

const markdown = workbook.exportMarkdown({}, { maxRows: 100 });
const fromBytes = await exportXlsxStructured(bytes); // no session, no clock
```

- Each sheet lists its stored cells in row-major order (formula cells and
  styled empty cells included, empty positions skipped) with value, formula,
  display text, number format and merge membership, plus merges, tables,
  hyperlinks, hidden row and column spans, and charts, pictures and shapes as
  placeholders with alt text. Defined names are listed read-only.
- Anchors are `{ sheet: { index, name }, a1 }` positions in the exported version
  (`anchorScope: "session"`) or bytes snapshot (`"snapshot"`); neither follows
  later row, column or sheet edits. Retained drawings carry a `sourcePart`
  provenance with the part's SHA-256.
- Formula results are the stored values (`calculation.policy: "asStored"`,
  `freshness: "unverified"`); cells mark results the file did not store as
  `missing` (`uncertain` once anything has calculated, or where it cannot be
  traced) and the last calculation's `cycle` and `limited` cells. Live exports
  reflect whatever calculation already ran, and the same version and options
  always export the same content; `exportXlsxStructured` reads bytes without
  recalculating.
- Hidden sheets, rows, columns and names are excluded unless requested
  (`includeHiddenSheets`, `includeHiddenRows`, `includeHiddenColumns`,
  `includeHiddenNames`), with a `hidden-content-excluded` diagnostic. A sheet
  whose visibility is unknown, such as one from a model handed in without its
  package, counts as hidden.
  Comments, rich-text runs, conditional formatting, pivot tables and unreadable
  charts are diagnosed, not exported.
- `maxCells` (default 100,000) and `maxBytes` (default 8 MiB) stop at a complete
  record with `truncated: true` and a `truncated` diagnostic; an absent cell is
  empty only before that point. Scope refusals (`invalid-scope`,
  `invalid-options`, `limit-exceeded`) return `{ ok: false, version, failure }`.
- Markdown renders one grid per sheet labelled with its A1 columns and row
  numbers (200 rows, 50 columns and 10,000 positions by default), using an
  escaped HTML table where merges need spans, and `<!-- xlsx-export:N -->`
  markers whose anchors come back in `anchors`. Document text is escaped and
  kept on one line, and hyperlinks become Markdown links only for `http`,
  `https` and `mailto` destinations. `renderXlsxMarkdown` renders content you
  already hold, refusing content that does not validate.

## Collaboration

Open a collaborative replica, then connect it to any reliable binary transport:

```ts
import { initWasm, openWorkbook } from "@betteroffice/xlsx";
import {
  CollaborationProvider,
  type CollaborationTransport,
} from "@betteroffice/xlsx/collaboration";

await initWasm();
const workbook = openWorkbook(bytes, { collaborative: true });
const transport: CollaborationTransport = createTransport();
const provider = new CollaborationProvider(workbook, transport, {
  user: { name: "Ada" }, // identity for this peer's selection flag
});
provider.connect();

function dispose() {
  provider.destroy();
  workbook.dispose();
}
```

The provider speaks the Yjs sync-v1 protocol used by y-websocket. WebSocket room
routing, authentication, WebRTC signaling, reconnection policy, and awareness
remain transport concerns; document updates flow directly between the connection
and the Rust/WASM Yrs replica without a second JavaScript `Y.Doc`.
After a close, the transport may reopen itself or the caller may invoke
`provider.connect()` for another connection attempt.
Call `provider.destroy()` before discarding its transport or workbook.

Collaborative sessions currently support cell content, formulas, styles, column
widths, and row heights. Structural edits and inverse-op undo are rejected until
they have stable axis identities and a Yrs-aware undo manager.

## Development

The generated `.wasm` binary is intentionally not committed. From the repository
root, install `wasm-pack` 0.15.0 and `binaryen`, then run
`bun run build:xlsx-wasm`. Package builds, tests, demo startup, and CI run this
step automatically.

[JavaScript guide](https://docs.betteroffice.dev/docs/javascript) ·
[Changelog](https://github.com/openooxml/betteroffice/blob/main/packages/xlsx/CHANGELOG.md) · Apache-2.0.

### Operation profiles

`editCellProfiled`, `applyOpsProfiled`, and `displayListProfiled` perform the same
operations as their unprofiled counterparts and return engine stage durations
in milliseconds. Edit results include `profile` (validate, apply, recalc, result);
display results contain `{ displayList, profile }` (build, encode). The normal
entry points do not read profiling clocks. See the
[corpus and browser tests](../../e2e/README.md) for usage and timing limits.
