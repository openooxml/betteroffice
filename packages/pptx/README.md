# @betteroffice/pptx

Framework-free core for the BetterOffice PPTX editor — the Rust parser, yrs deck
model, slide layout, and display-list engine compiled to WebAssembly, plus the
Canvas2D replay host.

```bash
bun add @betteroffice/pptx
```

Most apps want the turnkey React component in
[`@betteroffice/pptx-react`](https://www.npmjs.com/package/@betteroffice/pptx-react).
Use this package directly to build custom presentation chrome.

## Open and render a slide

```ts
import {
  initWasm,
  openPresentation,
  paintSlide,
  sizeCanvasForSlide,
} from '@betteroffice/pptx';

await initWasm();
const bytes = new Uint8Array(await file.arrayBuffer());
const deck = openPresentation(bytes, {
  fonts: [{ family: 'My Sans', bytes: fontBytes }],
});

const frame = deck.layoutSlide(0);
sizeCanvasForSlide(canvas, frame, devicePixelRatio);
await paintSlide(canvas.getContext('2d')!, frame, devicePixelRatio);
```

`initWasm()` with no argument only works where `fetch` of a same-origin URL works (browsers); Node and SSR must pass the wasm bytes.

All parsing, edits, collaboration state, text shaping, layout, hit-testing, and
display-list emission stay in Rust. The package decodes the typed boundary and
replays the resulting primitives on canvas. Font bytes are supplied by the host
and registered with the Rust shaper through `openPresentation`.

Beyond rendering, `PresentationHandle` covers editing: version-checked batches
(`readContent` / `findText` / `validateEdits` / `applyEdits`), read-only
structured export (`exportStructured` / `exportMarkdown`), text
(`insertText` / `deleteText` / `formatText` / `setParagraphAlignment`), slides
(`insertSlide` / `deleteSlide` / `moveSlide`), shapes
(`addTextBox` / `addShape` / `addPicture` / `moveShape` / `resizeShape` /
`setShapeRect`), paint order (`bringShapeToFront` / `sendShapeToBack` /
`bringShapeForward` / `sendShapeBackward`), comments
(`addComment` / `replyToComment` / `setCommentStatus` / `removeComment`),
`hitTest`, undo/redo, and `save()`, which serializes the deck back to `.pptx`
bytes with edits applied — untouched slides keep their exact source part bytes.
`addPicture` accepts common raster/vector MIME types (PNG, JPEG, GIF, BMP,
TIFF, WebP, SVG) up to 8 MiB, rejecting anything past either limit before it
reaches the deck.

## Agent proposals

Stage a group of edits, inspect it, then accept or reject it:

```ts
const slide = deck.snapshot().slides[0];
const shape = slide.shapes.find((shape) => shape.textStories.length > 0)!;
const story = shape.textStories[0];
const proposal = deck.propose('editor-agent', 'Make the title concise', [{
  type: 'replaceText',
  storyId: story.id,
  start: 0,
  end: story.paragraphs[0].runs.reduce((length, run) => length + run.text.length, 0),
  text: 'A clear title',
}]);

const preview = deck.previewProposal(proposal.id);
const proposedSlide = deck.layoutProposalSlide(proposal.id, 0);
const diff = deck.layoutProposalDiffSlide(proposal.id, 0);
await paintSlide(canvas.getContext('2d')!, diff.frame, devicePixelRatio, 1, {
  textChanges: diff.textChanges,
});
deck.acceptProposal(proposal.id);
deck.undo();
```

`listProposals()` returns agent attribution, notes, edits, before/after targets,
and current `staleTargets`. `previewProposal()` and `layoutProposalSlide()` replay
against the current deck without changing it. `rejectProposal(id)` removes a
pending group. Acceptance applies the whole group in one undo step and emits
one local update; an invalid edit prevents the entire group from being staged
or accepted.

`layoutProposalDiffSlide()` lays removed and inserted text out together, keeping
unchanged words as context and retaining fonts and emphasis. Pass its
`textChanges` to `paintSlide()` for red highlights and strikethrough on deletions,
and green highlights and underlines on insertions. Its snapshot and UTF-16
ranges describe a temporary review layout; use the live handle for editing and
the ordinary proposed layout for the result after acceptance. Review rendering
does not replace the live hit-test state or add markup to saved presentations.

Supported edits replace text within one paragraph, format text, align paragraphs,
set shape geometry/fill/stroke/adjustments, and replace speaker notes. Text offsets
use UTF-16 and shape coordinates use EMU. Text replacement inherits the style at
its start unless `style` is supplied. Edits run in order, so later text offsets
refer to the result of earlier edits in the group.

Changing a targeted shape or its notes makes acceptance throw `StaleProposalError`
with target IDs. Review a fresh preview before calling
`acceptProposal(id, { force: true })`; force still validates targets and ranges.
Unrelated peer edits survive acceptance and Undo. Up to 64 proposals, each with
1–256 edits, may be pending.

Pending proposals belong to this open session: they are excluded from PPTX
exports and collaboration updates. Accepted edits save and sync normally.
`isProposalsAvailable()` supports hosts that load an older WASM build.

## Version-checked edit batches

Read the deck with its session version, then apply a batch against that
version: every step commits in one transaction, one update and one undo step,
or the batch returns a typed refusal and nothing changes.

```ts
const read = deck.readContent();
if (!read.ok) throw new Error(read.failure.message);
const story = read.stories[0];
const within = { slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId };

const result = deck.applyEdits({
  expectVersion: read.version,
  steps: [
    { op: 'replaceText', target: { kind: 'search', within, text: 'Q3' }, text: 'Q4' },
    { op: 'setSlideNotes', target: { slideId: story.slideId }, text: 'Updated for Q4' },
  ],
});
if (!result.ok) console.warn(result.failure.code, result.failure.stepIndex);
```

A story reads as its paragraphs joined by `\n`. Offsets are story-local UTF-16
positions, and each entry of `paragraphs` gives a paragraph's span, its soft
line breaks (which also read as `\n`) and its field results. `findText()`
searches exactly, case-sensitively and within paragraphs; a `search` target
must match exactly once in its story, overlapping occurrences included. Every
target and guard resolves against `expectVersion`, so a later step never sees
an earlier step's offsets.

Steps insert, replace and delete text within one paragraph, format text, align
paragraphs, replace speaker notes, and set the rectangle, fill or outline of a
shape at the top of a slide (fill and outline on preset shapes). An `expect`
guard refuses its step unless the target still reads as expected. Receipts
locate each change in the final state. `validateEdits()` runs every check,
staging included, without changing anything. `history: 'none'` keeps a batch
out of undo history and existing undo and redo entries in place;
`source: 'agent'` records provenance only.

Refusals carry a `code` (`stale-version`, `missing-target`, `ambiguous-target`,
`content-mismatch`, `overlapping-steps`, `unsupported`, `invalid-step`,
`limit-exceeded`), the failing `stepIndex` and the target; malformed requests
throw, and so do NaN or infinite numbers. Current limits: text steps leave
fields and soft line breaks whole, and a batch refuses when saving could turn a
field into plain text, which it can rule out only while every field is non-empty
and sits in the paragraph's unchanged leading or trailing text; a paragraph's alignment and its
text cannot change in one batch; slides, shapes and paragraphs are neither
created nor removed; a batch holds at most 128 steps and 1,048,576 inserted
UTF-16 units; requests hold at most 16 MiB of JSON and reads and searches return
at most 64 MiB, searches marking the cut with `truncated`. Slide, shape, story and paragraph
ids anchor targets within one session only, and versions from one session
never match another.

## Structured export

`exportStructured()` returns the committed deck as JSON with the version it was
read at, and `exportMarkdown()` renders that same read as Markdown. Neither
flushes editor input or changes anything. `exportPptxStructured(bytes)`,
`exportPptxMarkdown(bytes)` and `renderPptxMarkdown(content)` do the same
headless, with anchors that address the returned snapshot only.

```ts
const read = deck.exportStructured({ includeNotes: true });
if (!read.ok) throw new Error(read.failure.message);
for (const slide of read.content.slides) {
  for (const shape of slide.shapes) {
    for (const paragraph of shape.stories.flatMap((story) => story.paragraphs)) {
      console.log(slide.index, paragraph.list?.kind, paragraph.anchor);
    }
  }
}

const { markdown, anchors } = await exportPptxMarkdown(bytes);
```

Slides come in deck order with their original index; shapes follow the current
shape tree depth first, a group's descendants at the group's position and table
cells row by row. This is the authored order, not a reading order inferred from
geometry. Paragraphs carry their level, effective alignment, authored
`bulletJson` and the resolved list marker (inherited from the layout, master
and text styles and numbered as the renderer numbers them); runs carry
formatting marks, links, soft line breaks, fields with their cached result and
zero-width placeholders for inline content such as equations. Tables keep their
grid, spans and merge continuations with each cell's current story. Pictures,
video, audio, charts, SmartArt, embedded objects and shape-tree elements the
deck model does not hold become placeholders with their alternative text and
relationships; their data is never exported. Layout and master content is not
exported, and an empty placeholder never shows its prompt text.

Every record carries an anchor: `text` ranges use the story offsets of
`readContent()`, `notes` and `comment` ranges index their plain text, and
records seeded from the file carry `provenance` (part, SHA-256, element path,
`sldId`, `cNvPr` id). Session anchors belong to the returned version and do not
survive save and reopen. A collaboration session opened from an update seeded by
an older release, without its source file, may not know which slides are hidden:
those slides are exported with `hidden: null` and a `visibility-unknown`
diagnostic, and reopening the session with its source file restores their
visibility. Hidden slides and shapes, speaker notes (plain text)
and comments are excluded unless `includeHiddenSlides`, `includeHiddenShapes`,
`includeNotes` or `includeComments` asks for them, and `includeFormatting:
false` drops marks. Everything left out or not represented is listed in
`diagnostics`. `maxBlocks` (10,000 by default; slides, shapes, paragraphs,
notes and comments count) and `maxBytes` (8 MiB of compact JSON) stop the
export at a whole record and set `truncated`. Unusable limits come back as an
`invalid-options` or `limit-exceeded` refusal (`PptxExportError` for the
headless functions); malformed options throw.

Markdown renders slide headings, title placeholders as headings, paragraphs
with literal list markers, pipe or entity-escaped HTML tables, object
placeholders, notes and comments, with a `<!-- pptx-export:N -->` marker before
each block that `anchors` maps back to its source.

## Comments

PowerPoint has two comment systems and a file only ever uses one: `legacy`
reads in every version of PowerPoint plus LibreOffice and Google Slides, while
`modern` carries replies and a resolved state but only shows in PowerPoint 365.
A deck commits to one at its first comment, so `setCommentFlavor` only works
while `comments()` is empty, and `replyToComment` / `setCommentStatus` throw on
a legacy deck.

Saving patches existing comment XML, preserving identities, formatting, task
metadata, anchors and unknown fields. Unedited parts remain byte-identical.
Removing a thread removes its known replies; a concurrent reply survives as a
new root if its parent was deleted.

```ts
deck.addComment(deck.snapshot().slides[0].id, {
  author: 'Ada Lovelace',
  initials: 'AL',
  text: 'Tighten this claim.',
  created: new Date().toISOString(),
  xEmu: 1_828_800,
  yEmu: 914_400,
});
```

Positions are EMU, like every other coordinate. `created` is supplied by the
caller rather than read from a clock, so peers replaying the same edits
converge.

## Collaboration

`PresentationHandle` is a collaboration replica. Pair it with
`CollaborationProvider` and a transport implementing the small
`CollaborationTransport` interface. The provider speaks the standard Yjs sync-v1
wire protocol, performs state-vector handshakes, forwards only local updates,
and bounds frames and pending backpressure bytes.

```ts
import { CollaborationProvider } from '@betteroffice/pptx';

const provider = new CollaborationProvider(deck, transport, {
  user: { name: "Ada" }, // identity for this peer's presence chip
});
deck.onUpdate((_update, origin) => {
  if (origin === 'remote') repaint();
});
provider.connect();
```

## Development

The generated `.wasm` binary is intentionally not committed. From the repository
root, install `wasm-pack` 0.15.0 and `binaryen`, then run
`bun scripts/build-pptx-wasm.ts`.
Package builds copy the binary into `dist/generated`.

[JavaScript guide](https://docs.betteroffice.dev/docs/javascript) ·
[Changelog](https://github.com/openooxml/betteroffice/blob/main/packages/pptx/CHANGELOG.md) · Apache-2.0.

### Operation profiles

`layoutSlideProfiled` returns `{ layout, profile }` with scope, layout, and
serialization durations in milliseconds. Profiled mutation methods such as
`insertTextProfiled` and `addTextBoxProfiled` return `{ receipt, profile }`;
`undoProfiled` reports undo, snapshot, and serialization timings. These methods
share the normal operations and add stage timing. See the
[corpus and browser tests](../../e2e/README.md) for usage and timing limits.

## Host undo and comment controls

`handle.setUndoCaptureMode('manual')` groups tracked local edits across pauses
and operation types until `handle.addUndoBoundary()`. The getter
`handle.undoCaptureMode()` returns `'auto'` or `'manual'`. Changing modes closes
the current group and preserves history; setting the same mode is a no-op.
Auto preserves the existing policy (500 ms capture on native targets, separate
transactions in the browser). Remote and agent origins remain outside local
undo. These controls group history; they do not make edits atomic.

`handle.setCommentPosition(commentId, { xEmu, yEmu })` moves an existing root
comment on its current slide, preserving identity, author, text, replies, and
resolution state through collaboration, undo/redo, and export. Coordinates must
be safe integers in EMU. Replies share their root's position. Unknown IDs,
reply IDs, and invalid coordinates fail before mutation. Legacy PowerPoint
comments quantize exported positions to master units (1/576 inch).
