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

### Reanchor an existing comment

`session.setCommentRanges(commentId, ranges)` replaces only the sticky anchors of
an existing comment. The id, author, date, body, reply relationship and resolution
state remain intact. Ranges use the same paragraph locations as `addComment`;
one range may span paragraphs, and separate ranges may address different stories.

Every range must be non-empty, ordered, and within existing paragraphs. An empty
list, unknown comment/story/paragraph, or invalid offset throws before any content
changes. The host must find the surviving text and supply its new range; the API
does not infer text matches after replacement.

If replacement removes all commented text, the host must explicitly choose a new
non-empty range or handle the comment's removal. Rejected reanchoring leaves the
comment unchanged and does not roll back a text replacement already performed.

Reanchoring joins the current local undo capture, so an immediate replacement and
reanchor can undo together. Comment edits participate in the session's history;
undoing comment changes conservatively invalidates all stories for saved anchors.

The undo manager retains comment item boundaries with their undo/redo entries so
anchors inside replaced text survive repeated history traversal with Yrs 0.27.
Discarding history releases its bookkeeping; undo/redo traverses a local snapshot
when those boundaries need restoration.
