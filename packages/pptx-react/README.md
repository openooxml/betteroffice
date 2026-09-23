# @betteroffice/pptx-react

React chrome for the BetterOffice PPTX editor — wraps
[`@betteroffice/pptx`](https://www.npmjs.com/package/@betteroffice/pptx) in a
slide canvas, slide strip, formatting toolbar, and keyboard editing surface.
Parsing, layout, and text shaping run in the core's Rust/WebAssembly engine;
slides are painted onto canvas.

<!-- TODO(author): add a screenshot/GIF here once hosted; an <img> with an unresolvable src renders broken on npm -->

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
for bold, italic, size, color, slides, text boxes, image insertion, and shape ordering.

Props: `file`, `fonts`, `collaboration`, `i18n`, `className`, `fileName`,
`onReady` (exposes the core `PresentationHandle`, a `refresh` callback for
host-driven edits, `refreshProposals`, and `save`), `onChange` (deck snapshots), `onError`, and
`onSave` (receives the saved bytes; without it, saving downloads the file).

## Host editing controls

`onSaveRequest` runs for toolbar and Ctrl/Cmd+S saves before serialization.
Return `true` to continue built-in saving; `false` or `void` handles or cancels
it. Promises are awaited and concurrent requests are coalesced. `onSave` still
receives the resulting bytes when built-in saving continues. A request waiting
on a replaced or closed document is discarded.

The API received by `onReady` exposes `flushPendingInput(): Promise<void>`.
Await it before inspecting or mutating the core from a host workflow, then call
`api.save()` for explicit serialization without re-entering `onSaveRequest`.
`save()` remains synchronous and rejects while asynchronous input is pending.
Flush rejects stale document handles, failed input, and unfinished pointer
gestures. Finish or cancel the gesture before retrying.

PPTX keyboard and notes edits are synchronous; flushing also waits for accepted
image imports. Reported asynchronous input failures continue to reject flushing
until the document is reopened.

`api.getPositionAtPoint(clientX, clientY)` returns a shape or text hit with
`slide` (1-based), `slideId`, and `shapeId`; text hits include `storyId` and the
UTF-16 `position`. It uses the current slide and canvas bounds without changing
focus or selection. Outside content, unavailable geometry, stale handles, and
proposal previews return `null`. Use `api.refresh()` after core mutations.

## What works today

- Slide rendering with Rust layout and text shaping, painted onto canvas
- Canvas interactions: shape selection, drag, and resize
- Text editing with caret and selection computed by the engine
- Slide management (add, delete), text boxes, image insertion, shape ordering, undo/redo
- Saving the deck back to `.pptx` — toolbar button, Ctrl/Cmd+S, or `api.save()`
- Localized UI via the `i18n` prop
  ([`@betteroffice/pptx-i18n`](https://www.npmjs.com/package/@betteroffice/pptx-i18n))
- Real-time collaboration with people or agents; the deck is a CRDT
- Live collaborator shape selections and presence chips, shown in each peer's color
- Agent proposal review directly on the slide canvas, with inline text diffs,
  old/new shape bounds, before/after previews, acceptance, rejection, and Undo

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

Pending edits appear directly on the slide canvas: deleted text is red and
struck through, inserted text is green and underlined, and moved/resized shapes
show their previous and proposed bounds. Speaker notes have their own inline
text diff. The canvas toolbar selects a proposal when several affect the slide
and provides acceptance, rejection, and access to review details. Acceptance is
available after the current diff has painted successfully.

The canvas is in review mode while showing a diff. **Edit slide** (or Escape)
returns to normal editing; **Show changes** restores the diff. This keeps
temporary review offsets separate from editable text. Saving, PNG export, and
presentation mode use the actual document. Changes from other users refresh the
review, and stale targets require the explicit review described below.

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

[JavaScript guide](https://docs.betteroffice.dev/docs/javascript) ·
[Changelog](https://github.com/openooxml/betteroffice/blob/main/packages/pptx-react/CHANGELOG.md) · Apache-2.0.
