# @betteroffice/xlsx

Framework-free core for the BetterOffice XLSX editor — the Rust engine (parse,
calc, render) compiled to WebAssembly, plus display-list, viewport, hit-test,
and accessibility helpers.

> **Early (`0.0.x`).** The core surfaces — opening/saving documents, the editor
> components, collaboration — are settling and unlikely to change shape. Smaller
> APIs may still move between releases; breaking changes are always listed in
> the changelog.

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

## AI agents / human-in-the-loop

An agent stages edits as a proposal instead of applying them; a human reviews
per-cell before/after previews and accepts or rejects. Accepting applies the
proposal as one undo step and recalcs dependents, and throws
`StaleProposalError` if the workbook drifted under the proposal since it was
staged.

```ts
const proposal = workbook.propose("copilot", "add totals", [
  { sheet: 0, row: 9, col: 2, input: "=SUM(C1:C9)" },
]);

workbook.listProposals(); // pending proposals, oldest first
workbook.acceptProposal(proposal.id); // or workbook.rejectProposal(proposal.id)
```

The React editor paints pending proposals as in-cell tracked-change ghosts with
an accept/reject panel. Guard with `isProposalsAvailable()` against cores built
without the feature.

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
const provider = new CollaborationProvider(workbook, transport);
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
root, install `wasm-pack` 0.15.0 and run `bun run build:xlsx-wasm`. Package builds,
tests, demo startup, and CI run this step automatically.

Docs: https://betteroffice.dev · Apache-2.0.
