# @betteroffice/pptx-react

React chrome for the BetterOffice PPTX editor — wraps
[`@betteroffice/pptx`](https://www.npmjs.com/package/@betteroffice/pptx) in a
slide canvas, slide strip, formatting toolbar, and keyboard editing surface.

> **Experimental (`0.0.x`).** The API is unstable and may change in any release.

```bash
bun add @betteroffice/pptx-react @betteroffice/pptx react react-dom
```

`react` and `react-dom` (18 or 19) are peer dependencies.

## Usage

```tsx
import { PptxEditor } from '@betteroffice/pptx-react';

export function Presentation({ file, fontBytes }: {
  file: Uint8Array;
  fontBytes: Uint8Array;
}) {
  return (
    <div style={{ height: 720 }}>
      <PptxEditor
        file={file}
        fonts={[{ family: 'My Sans', bytes: fontBytes }]}
        onChange={(deck) => console.log(deck.slides.length)}
      />
    </div>
  );
}
```

Click text to place the Rust-computed caret, type to edit the yrs story and
trigger Rust reflow, or use the toolbar for bold, italic, size, color, slides,
and text boxes. `onReady` exposes the core handle and a `refresh` callback for
host-driven edits.

## Collaboration

Pass a `collaboration` prop to co-edit a deck live. `onReplica` hands you the
session; drive it with a `CollaborationProvider` over any transport. Every peer
must boot from the same `initialUpdate` seed.

```tsx
import { CollaborationProvider } from '@betteroffice/pptx/collaboration';

<PptxEditor
  file={file}
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

Docs: https://betteroffice.dev · Apache-2.0.
