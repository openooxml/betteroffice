# @betteroffice/docx-react

React chrome for the BetterOffice DOCX editor — wraps
[`@betteroffice/docx`](https://www.npmjs.com/package/@betteroffice/docx) in a
drop-in `<DocxEditor>` component with the toolbar, ruler, selection, keyboard,
comments, and tracked-changes UI wired up. Rendering and layout run in the
core's Rust/WebAssembly engine; pages are painted onto canvas.

<!-- TODO(author): add a screenshot/GIF here once hosted; an <img> with an unresolvable src renders broken on npm -->

> **Early (`0.0.x`).** The core surfaces — opening/saving documents, the editor
> components, collaboration — are settling and unlikely to change shape. Smaller
> APIs may still move between releases; breaking changes are always listed in
> the changelog.

```bash
bun add @betteroffice/docx-react @betteroffice/docx react react-dom
```

`react` and `react-dom` (18 or 19) are peer dependencies.

## View a document

```tsx
import { DocxEditor } from '@betteroffice/docx-react';
import '@betteroffice/docx-react/styles.css';

<DocxEditor documentBuffer={buffer} mode="viewing" />;
```

`documentBuffer` accepts an `ArrayBuffer`, `Uint8Array`, `Blob`, or `File`.

## Edit a document

```tsx
import { useState } from 'react';
import { DocxEditor } from '@betteroffice/docx-react';
import '@betteroffice/docx-react/styles.css';

export function App() {
  const [file, setFile] = useState<ArrayBuffer>();

  return (
    <>
      <input
        type="file"
        accept=".docx"
        onChange={async (e) => {
          const f = e.target.files?.[0];
          if (f) setFile(await f.arrayBuffer());
        }}
      />
      <DocxEditor
        documentBuffer={file}
        onSave={(bytes) => console.log(`saved ${bytes.byteLength} bytes`)}
      />
    </>
  );
}
```

Without `onSave`, File > Save downloads the edited bytes.

Key props: `documentBuffer` (or a parsed `document`), `onSave`, `onChange`,
`author`, `mode` (`editing` / `suggesting` / `viewing`), `showToolbar`,
`showRuler`, `showZoomControl`, `i18n`, `measurementFontProvider`. The `ref`
exposes the full editor API (selection, formatting, find/replace, comments,
revisions).

## What works today

- Editing with Word-faithful pagination; layout runs in Rust, never in the DOM
- Suggesting mode: tracked changes with accept/reject review UI
- Comment threads with replies and resolution, controllable from the host
- Find and replace, headers and footers, footnotes, images, tables
- Zoom control and ruler
- Localized UI via the `i18n` prop
  ([`@betteroffice/docx-i18n`](https://www.npmjs.com/package/@betteroffice/docx-i18n))
- Real-time collaboration with people or agents; the document is a CRDT
- Live collaborator cursors and selections, shown in each peer's color

For Word-accurate metrics install `@betteroffice/fonts` and hand it to the
engine once at module scope — `configureDefaultFonts({ fonts })`, re-exported
here — or pass your own provider through the `measurementFontProvider` prop.
Without either, measurement falls back to the browser.

## Collaboration

The document is a CRDT — pass a `collaboration` prop and wire a transport to
co-edit with other people or an agent. `onReplica` hands you the session as a
`CollaborationReplica`; drive it with a `CollaborationProvider` over any
transport (a WebSocket relay, etc.). Every window must boot from the same shared
state, so pass identical `initialUpdate` bytes to each peer.

```tsx
import { CollaborationProvider } from '@betteroffice/docx/collaboration';

<DocxEditor
  documentBuffer={file}
  collaboration={{
    clientId,
    initialUpdate: sharedSeed, // identical bytes for every peer
    onReplica: (replica) => {
      if (!replica) return;
      const provider = new CollaborationProvider(replica, transport, {
        user: { name: "Ada" }, // shown on this peer's remote caret
      });
      provider.connect();
    },
  }}
/>;
```

## Framework notes

Import `@betteroffice/docx-react/styles.css` once (in a bundler entry or, under
Next.js, at the page/layout level — CSS imported inside a `next/dynamic`
component does not attach in production builds). The editor is browser-only
(canvas, wasm, workers); under Next.js load it with `next/dynamic` and
`ssr: false`.

Docs: https://betteroffice.dev · Apache-2.0.
