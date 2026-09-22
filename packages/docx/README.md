# @betteroffice/docx

Framework-free core for the BetterOffice DOCX editor — the Rust engine (OOXML
parse/serialize, CRDT editing core, text shaping, pagination) compiled to
WebAssembly, plus the display-list, canvas-render, geometry, and accessibility
helpers the adapters build on. Layout never touches the DOM: the engine
measures every line and pages are replayed onto canvas.

```bash
bun add @betteroffice/docx
```

Most apps want the turnkey React component in
[`@betteroffice/docx-react`](https://www.npmjs.com/package/@betteroffice/docx-react).
Reach for this package directly for headless parsing/serialization or when
building a custom adapter.

## Parse and save

A full round trip: open a `.docx`, inspect the typed model, write bytes back.

```ts
import { readFile, writeFile } from 'node:fs/promises';
import { parseDocx, repackDocx } from '@betteroffice/docx/docx';

const document = await parseDocx(await readFile('contract.docx'));
// document.package: body, styles, numbering, theme, media, headers/footers

const bytes = await repackDocx(document);
await writeFile('contract-out.docx', Buffer.from(bytes));
```

`repackDocx` round-trips against the original buffer so untouched parts are
preserved; use `createDocx` for documents built from scratch.

The engine ships as four wasm assets (container, parser, layout, editing core)
in `dist/generated/`. Browsers fetch them lazily behind the async entry points
(`parseDocx`, save, the layout engine, `createYrsSession`); Node and Bun read
them from disk synchronously on first use. No manual init call is required.

## Collaboration

Connect the editor's Yrs replica to any reliable binary transport:

```ts
import { CollaborationProvider } from '@betteroffice/docx/collaboration';

const provider = new CollaborationProvider(replica, createTransport(), {
  user: { name: "Ada" }, // identity for this peer's remote caret
});
provider.connect();
```

`replica` can be a direct `YrsSession`, the worker-aware adapter returned by
`createWorkerCollaborationReplica`, or the value published by the React
editor's `collaboration.onReplica` callback. The provider speaks Yjs sync-v1;
room routing, authentication, awareness, and reconnection policy remain
transport concerns. Pass a persisted Yrs update as `collaboration.initialUpdate`
when a React editor joins an existing room so it hydrates the shared history
instead of independently importing the same DOCX.

## Development

The generated `.wasm` binaries are intentionally not committed. From the
repository root, install `wasm-pack` 0.15.0 and `binaryen`, then run
`bun run build:docx-wasm`.
Package builds, demo startup, and CI run this step automatically.

[JavaScript guide](https://docs.betteroffice.dev/docs/javascript) ·
[Changelog](https://github.com/openooxml/betteroffice/blob/main/packages/docx/CHANGELOG.md) · Apache-2.0.

### Separate host undo actions

Call `session.stopUndoCapture()` between independent programmatic actions to keep
them in separate undo steps, even within the 500 ms typing capture window. It
closes the current capture without disabling history, adding an empty step, or
clearing redo. Repeated calls and calls before `beginUndoCapture()` are safe.

```ts
session.stopUndoCapture();
session.insertText(location, 'Prefixo ');
session.stopUndoCapture();
session.deleteRange(markerRange);
session.stopUndoCapture();
```

Place boundaries before and after a host action to isolate it from surrounding
typing. Normal typing retains its existing coalescence behavior; this API does not
make asynchronous edits or edits across stories one atomic transaction.
