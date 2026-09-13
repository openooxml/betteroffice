# @betteroffice/pptx-react

React chrome for the BetterOffice PPTX editor — wraps
[`@betteroffice/pptx`](https://www.npmjs.com/package/@betteroffice/pptx) in a
slide canvas, slide strip, formatting toolbar, and keyboard editing surface.
Parsing, layout, and text shaping run in the core's Rust/WebAssembly engine;
slides are painted onto canvas.

<!-- TODO(author): add a screenshot/GIF here once hosted; an <img> with an unresolvable src renders broken on npm -->

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
with [`@betteroffice/fonts`](https://www.npmjs.com/package/@betteroffice/fonts)
for metric-compatible open faces.

## Edit

Click text to place the Rust-computed caret, type to edit the yrs story and
trigger Rust reflow, drag or resize shapes on the canvas, or use the toolbar
for bold, italic, size, color, slides, and text boxes.

Props: `file`, `fonts`, `collaboration`, `i18n`, `className`, `fileName`,
`onReady` (exposes the core `PresentationHandle`, a `refresh` callback for
host-driven edits, `refreshProposals`, and `save`), `onChange` (deck snapshots), `onError`, and
`onSave` (receives the saved bytes; without it, saving downloads the file).

## What works today

- Slide rendering with Rust layout and text shaping, painted onto canvas
- Canvas interactions: shape selection, drag, and resize
- Text editing with caret and selection computed by the engine
- Slide management (add, delete), text boxes, undo/redo
- Saving the deck back to `.pptx` — toolbar button, Ctrl/Cmd+S, or `api.save()`
- Localized UI via the `i18n` prop
  ([`@betteroffice/pptx-i18n`](https://www.npmjs.com/package/@betteroffice/pptx-i18n))
- Real-time collaboration with people or agents; the deck is a CRDT
- Live collaborator shape selections and presence chips, shown in each peer's color
- Agent proposal review with target navigation, rendered before/after slides,
  acceptance, rejection, stale-target review, and Undo

## Review agent proposals

Use the editor API supplied to `onReady` to stage edits from your agent, then
refresh the pending list:

```ts
api.handle.propose('editor-agent', 'Clarify the speaker notes', [{
  type: 'setSlideNotes',
  slideId: api.handle.snapshot().slides[0].id,
  text: 'Explain the customer outcome before the implementation details.',
}]);
api.refreshProposals();
```

The **Agent proposals** button shows the pending count. Each group shows its
author, rationale, targets, and text changes. **Preview** opens current and
proposed slide renderings, with a target selector for groups spanning multiple
shapes or slides. Notes changes also appear as text because notes are outside
the slide canvas.

Accepting a group updates the editor, calls `onChange`, and creates one Undo
step. Rejecting it leaves the deck untouched. If a target changed, preview its
current state before choosing **Apply updated proposal**. A further target
change detected at that click refreshes the preview for another review.

Pending proposals are session-local and disappear when the deck closes. Only
accepted edits are saved and synchronized. Existing host-driven edits may keep
using `api.refresh()`, which also refreshes the proposal list.

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
      const provider = new CollaborationProvider(replica, transport, {
        user: { name: "Ada" }, // shown on this peer's presence chip
      });
      provider.connect();
    },
  }}
/>;
```

## Framework notes

The editor is browser-only (canvas, wasm); under Next.js load it with
`next/dynamic` and `ssr: false`.

Docs: https://betteroffice.dev · Apache-2.0.
