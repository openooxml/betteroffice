# @betteroffice/pptx

Framework-free PPTX parsing, editing, rendering, and saving through a Rust/WebAssembly engine.
Use it to build a presentation viewer or editor with your own interface.

## Install

```bash
bun add @betteroffice/pptx
```

No peer dependencies.

## Usage

Render the first slide of a presentation onto a browser canvas:

```ts
import {
  initWasm,
  openPresentation,
  paintSlide,
  sizeCanvasForSlide,
} from '@betteroffice/pptx';

export async function renderFirstSlide(
  bytes: Uint8Array,
  fontBytes: Uint8Array,
  canvas: HTMLCanvasElement,
) {
  await initWasm();
  const deck = openPresentation(bytes, {
    fonts: [{ family: 'My Sans', bytes: fontBytes }],
  });
  try {
    const frame = deck.layoutSlide(0);
    const dpr = window.devicePixelRatio;
    sizeCanvasForSlide(canvas, frame, dpr);
    const context = canvas.getContext('2d');
    if (!context) throw new Error('Canvas 2D is unavailable');
    await paintSlide(context, frame, dpr);
  } finally {
    deck.dispose();
  }
}
```

`fontBytes` holds a TTF or OTF file, for example fetched from your server or loaded with
`loadBundledFontBytes` from [`@betteroffice/fonts`](https://www.npmjs.com/package/@betteroffice/fonts). Slides containing images also need an image resolver, described
below. The React editor in [`@betteroffice/pptx-react`](https://www.npmjs.com/package/@betteroffice/pptx-react)
manages rendering and user input.

## Entry points

All entry points below are exported from `@betteroffice/pptx`.

| entry point | description |
| --- | --- |
| `initWasm(input?)` | Initialize WebAssembly before opening or inspecting a file. Browsers can omit the input; Node hosts can pass WASM bytes. |
| `openPresentation(bytes, options?)` | Open a synchronous `PresentationHandle` for reads, edits, layout, and saving. |
| `inspectPresentation(bytes, options?)` | Inspect parsed content without opening an editing handle; `includeMedia: false` omits media. |
| `sizeCanvasForSlide(canvas, frame, dpr, scale?)` | Set canvas backing dimensions and CSS size for a slide. |
| `paintSlide(context, frame, dpr?, scale?, options?)` | Paint a display list asynchronously onto a canvas. |
| `slideToPng(frame, options?)` | Return a PNG `Blob` asynchronously using a browser canvas backend. |
| `exportPptxStructured(bytes, options?)` | Export structured content asynchronously without fonts or a DOM. |
| `exportPptxMarkdown(bytes, options?)` | Export Markdown and source anchors asynchronously. |
| `renderPptxMarkdown(content, options?)` | Render structured content as Markdown asynchronously. |

`openPresentationSession(bytes, options?)` from `@betteroffice/pptx/worker` opens an experimental asynchronous
worker session.

`openPresentation` accepts `fonts`, `fallbackFonts`, `clientId`, and `initialUpdate`. Fallback fonts
supply glyphs missing from a run's own face. When opening from a collaboration update, retain the
source file bytes used to seed it so the handle can save.

The core's directly owned handle is synchronous. The React editor runs its engine in a Web Worker
by default and exposes promise-returning reads and exports.

## Layout and rendering

`layoutSlide(index)` uses a zero-based slide index. `snapshot()` returns the deck model, and
`layouts()` lists layout IDs and names for `insertSlide`. `duplicateSlide` copies slide content and
speaker notes, excluding comments. Shape coordinates use EMU (English Metric Units, 914,400 per inch); text offsets use
UTF-16 code units.

For embedded images, pass `paintSlide` an options object with `resolveImage`, a function mapping an
asset ID to a canvas image source or a promise for one. It may return `null` for an unavailable image.
`slideToPng` accepts the same resolver and an optional `scale`.

## Saving and ownership

`save()` returns PPTX bytes with edits applied; untouched slides retain their source part bytes.
Call `dispose()` when finished with a handle. `fork()` creates an independently owned copy that must
also be disposed.

## Agent proposals

`propose(agentId, note, edits)` stages edits without changing the committed deck. Use `listProposals`,
`previewProposal`, and `layoutProposalSlide` to inspect them, then `acceptProposal` or `rejectProposal`.
Acceptance applies the group in one undo step; pending proposals stay in the open session.

`layoutProposalDiffSlide` returns a review frame and `textChanges`; pass both to `paintSlide` to show
insertions and deletions. Review frames contain temporary text offsets and are for display only.
If a target changed, acceptance throws `StaleProposalError`; inspect a fresh preview before choosing
`acceptProposal(id, { force: true })`.

## Version-checked edit batches

`readContent`, `findText`, `validateEdits`, and `applyEdits` support edits against a session version.
See [Edit PPTX in version-checked batches](https://docs.betteroffice.dev/docs/javascript#edit-pptx-in-version-checked-batches)
for request shapes, atomicity, refusals, and limits.

## Structured export

`exportStructured` and `exportMarkdown` read committed state; headless exports accept file bytes.
See [PPTX structured export](https://docs.betteroffice.dev/docs/javascript#pptx-structured-export)
for options, anchors, diagnostics, and truncation.

## Comments and history

`comments()` lists comments. `setCommentFlavor` selects `legacy` or `modern` only while that list is
empty. `replyToComment` supports both flavors; `setCommentStatus` requires modern comments. `addComment` takes the author's
name, text, and caller-supplied `created` timestamp. `setCommentPosition` moves a root comment using
safe integer EMU coordinates; replies share their root's position.

`setUndoCaptureMode('manual')` groups tracked local edits until `addUndoBoundary()`;
`undoCaptureMode()` reads the current mode. `undo()` and `redo()` navigate history.
`anchorCaret(storyId, index)` creates an anchor that follows edits; `resolveCaretAnchor` returns its
current offset, or `null` if its story no longer exists.

## Collaboration

`PresentationHandle` implements the replica interface used by `CollaborationProvider`
from `@betteroffice/pptx/collaboration`.
See the [collaboration guide](https://docs.betteroffice.dev/docs/collaboration) for transports,
shared seeds, presence, and connection lifecycle.

## Operation profiles

`layoutSlideProfiled`, `insertTextProfiled`, and `undoProfiled` expose engine stage timings.
See [Operation profiles](https://docs.betteroffice.dev/docs/javascript#operation-profiles) for result
shapes and measurement guidance.

## Documentation

- [JavaScript guide](https://docs.betteroffice.dev/docs/javascript)
- [Collaboration guide](https://docs.betteroffice.dev/docs/collaboration)
- [Package guide](https://docs.betteroffice.dev/docs/packages)
- [React editor: @betteroffice/pptx-react](https://www.npmjs.com/package/@betteroffice/pptx-react)
- [Fonts: @betteroffice/fonts](https://www.npmjs.com/package/@betteroffice/fonts)
- [Repository](https://github.com/openooxml/betteroffice)
