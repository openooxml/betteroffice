# @betteroffice/docx

Framework-free core for the BetterOffice DOCX editor — the Rust engine (OOXML
parse/serialize, CRDT editing core, text shaping, pagination) compiled to
WebAssembly, plus the display-list, canvas-render, geometry, and accessibility
helpers the adapters build on. Layout never touches the DOM: the engine
measures every line and pages are replayed onto canvas.

> **Experimental (`0.0.x`).** The API is unstable and may change in any release.

```bash
bun add @betteroffice/docx
```

Most apps want the turnkey React component in
[`@betteroffice/docx-react`](https://www.npmjs.com/package/@betteroffice/docx-react).
Reach for this package directly for headless parsing/serialization or when
building a custom adapter.

## Usage

```ts
import { readFile } from 'node:fs/promises';
import { parseDocx } from '@betteroffice/docx/docx';

const buffer = await readFile('contract.docx');
const document = await parseDocx(buffer);
```

The engine ships as four wasm assets (container, parser, layout, editing core)
in `dist/generated/`. Browsers fetch them lazily behind the async entry points
(`parseDocx`, save, the layout engine, `createYrsSession`); Node and Bun read
them from disk synchronously on first use. No manual init call is required.

## Collaboration

Connect the editor's Yrs replica to any reliable binary transport:

```ts
import { CollaborationProvider } from '@betteroffice/docx/collaboration';

const provider = new CollaborationProvider(replica, createTransport());
provider.connect();
```

`replica` can be a direct `YrsSession`, the worker-aware adapter returned by
`createWorkerCollaborationReplica`, or the value published by the React
editor's `collaboration.onReplica` callback. The provider speaks Yjs sync-v1;
room routing, authentication, awareness, and reconnection policy remain
transport concerns.

## Development

The generated `.wasm` binaries are intentionally not committed. From the
repository root, install `wasm-pack` 0.15.0 and run `bun run build:docx-wasm`.
Package builds, demo startup, and CI run this step automatically.

Docs: https://betteroffice.dev · Apache-2.0.
