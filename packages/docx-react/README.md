# @betteroffice/docx-react

React chrome for the BetterOffice DOCX editor — wraps
[`@betteroffice/docx`](https://www.npmjs.com/package/@betteroffice/docx) in a
drop-in `<DocxEditor>` component with the toolbar, ruler, selection, keyboard,
comments, and tracked-changes UI wired up. Rendering and layout run in the
core's Rust/WebAssembly engine; pages are painted onto canvas.

> **Experimental (`0.0.x`).** The API is unstable and may change in any release.

```bash
bun add @betteroffice/docx-react @betteroffice/docx react react-dom
```

`react` and `react-dom` (18 or 19) are peer dependencies.

## Usage

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
        onSave={(bytes) => {
          // edited .docx bytes; without onSave, File > Save downloads them
          console.log(`saved ${bytes.byteLength} bytes`);
        }}
      />
    </>
  );
}
```

Import `@betteroffice/docx-react/styles.css` once (in a bundler entry or, under
Next.js, at the page/layout level — CSS imported inside a `next/dynamic`
component does not attach in production builds).

Bundled metric-compatible fonts ship in-repo, and the `measurementFontProvider`
prop accepts a custom provider for Word-accurate metrics.

Key props: `documentBuffer` (or a parsed `document`), `onSave`, `onChange`,
`author`, `mode` (`editing` / `suggesting` / `viewing`), `showToolbar`,
`showRuler`, `showZoomControl`, `measurementFontProvider`. The `ref` exposes
the full editor API (selection, formatting, find/replace, comments, revisions).

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
      const provider = new CollaborationProvider(replica, transport);
      provider.connect();
    },
  }}
/>;
```

Docs: https://betteroffice.dev · Apache-2.0.
