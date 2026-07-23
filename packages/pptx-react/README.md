# @betteroffice/pptx-react

React chrome for the BetterOffice PPTX editor — wraps
[`@betteroffice/pptx`](https://www.npmjs.com/package/@betteroffice/pptx) in a
slide canvas, slide strip, formatting toolbar, and keyboard editing surface.
Parsing, layout, and text shaping run in the core's Rust/WebAssembly engine;
slides are painted onto canvas.

<!-- TODO(author): screenshot/GIF — editor with slide strip / collab cursors -->
<img src="https://betteroffice.dev/readme/pptx-editor.png" alt="PptxEditor with slide strip and formatting toolbar" width="720" />

> **Early (`0.0.x`).** The core surfaces — opening/saving documents, the editor
> components, collaboration — are settling and unlikely to change shape. Smaller
> APIs may still move between releases; breaking changes are always listed in
> the changelog.

```bash
bun add @betteroffice/pptx-react @betteroffice/pptx react react-dom
```

`react` and `react-dom` (18 or 19) are peer dependencies.

## Render a presentation

```tsx
import { PptxEditor } from '@betteroffice/pptx-react';

export function Presentation({ file, fontBytes }: {
  file: Uint8Array;
  fontBytes: Uint8Array;
}) {
  return (
    <div style={{ height: 720 }}>
      <PptxEditor file={file} fonts={[{ family: 'My Sans', bytes: fontBytes }]} />
    </div>
  );
}
```

Font bytes are supplied by the host and registered with the Rust shaper; pair
with [`@betteroffice/docx-fonts`](https://www.npmjs.com/package/@betteroffice/docx-fonts)
for metric-compatible open faces.

## Edit

Click text to place the Rust-computed caret, type to edit the yrs story and
trigger Rust reflow, drag or resize shapes on the canvas, or use the toolbar
for bold, italic, size, color, slides, and text boxes.

Props: `file`, `fonts`, `collaboration`, `i18n`, `className`, `onReady`
(exposes the core `PresentationHandle` and a `refresh` callback for host-driven
edits), `onChange` (deck snapshots), `onError`.

## What works today

- Slide rendering with Rust layout and text shaping, painted onto canvas
- Canvas interactions: shape selection, drag, and resize
- Text editing with caret and selection computed by the engine
- Slide management (add, delete), text boxes, undo/redo
- Localized UI via the `i18n` prop
  ([`@betteroffice/pptx-i18n`](https://www.npmjs.com/package/@betteroffice/pptx-i18n))
- Real-time collaboration with people or agents; the deck is a CRDT
- Live collaborator cursors are landing in the next release.

## Collaboration

Pass a `collaboration` prop to co-edit a deck live. `onReplica` hands you the
session; drive it with a `CollaborationProvider` over any transport. Every peer
must boot from the same `initialUpdate` seed.

```tsx
import { CollaborationProvider } from '@betteroffice/pptx';

<PptxEditor
  file={file}
  fonts={fonts}
  collaboration={{
    clientId,
    initialUpdate: sharedSeed,
    onReplica: (replica) => {
      if (!replica) return;
      const provider = new CollaborationProvider(replica, transport);
      provider.connect();
    },
  }}
/>;
```

## Framework notes

The editor is browser-only (canvas, wasm); under Next.js load it with
`next/dynamic` and `ssr: false`.

Docs: https://betteroffice.dev · Apache-2.0.
