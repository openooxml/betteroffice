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

### Undo capture modes and boundaries

`session.setUndoCaptureMode(mode)` selects how tracked local transactions are
combined into undo steps. `session.undoCaptureMode()` returns the current mode.

| Mode | Grouping |
| --- | --- |
| `auto` (default) | Edits within 500 ms coalesce; switching stories closes the group. |
| `manual` | Edits coalesce across pauses and stories until an explicit boundary. |

`session.addUndoBoundary()` closes the current group in any mode. Changing modes
also closes the group; setting the same mode again does not. Both operations
preserve undo/redo history and are safe before capture starts. Repeated boundaries
create no empty undo steps. Remote transactions remain outside local undo history.

```ts
session.setUndoCaptureMode('manual');
session.insertText(firstLocation, 'Prefixo ');
session.insertText(secondLocation, 'Suffix');
session.addUndoBoundary();
session.deleteRange(markerRange);
session.addUndoBoundary();
session.setUndoCaptureMode('auto');
```

The two insertions undo together; marker deletion is a separate step. Boundaries
are also available in automatic mode when a host action needs to be isolated from
surrounding typing. Undo/redo close capture as usual. Manual mode controls history
grouping; it does not defer updates, flush pending input, or provide atomic execution.
The host must close manual groups so later unrelated edits do not join them.

### Paragraph identities

A session addresses paragraphs by session keys (`YrsLoc.paraId`); Word stores its
own paragraph ID (`w14:paraId`) in the file. A session key is never saved; a session
anchor resolves on every replica of one collaborative session, and each seeding open
(`openDocx`, `seedFromDocx`, `documentToYrs`) starts a new one unless given a fixed
`generation`, as a deterministic shared seed needs. Source IDs are kept as authored, and every
paragraph authored in the session gets a fresh, valid ID that avoids every ID the
package already uses.

`saveYrsDocx(session)` saves a session opened from DOCX bytes and returns each saved
paragraph's persisted anchor, qualified by its package part:

```ts
import { createYrsSession, saveYrsDocx } from '@betteroffice/docx/yrs';

const { secondParaId } = session.splitParagraph({ story: 'body', paraId, offset: 12 });
const saved = await saveYrsDocx(session);
const anchor = saved.paragraphs.find((p) => p.session.paraId === secondParaId)!.persisted;

const reopened = await createYrsSession();
reopened.openDocx(saved.bytes, true);
reopened.resolveParagraphAnchor(anchor); // { status: 'found', anchor: { kind: 'session', … } }
```

`saved.conflicts` lists saved paragraphs whose ID the live session reassigned while the
save ran, such as by a duplicate repair after a remote update; their anchors find them in
the saved bytes only. The React editor's Save writes the same IDs and records them as
saved; only `saveYrsDocx` returns the anchors and keeps unchanged parts as source bytes.
Source paragraphs without an ID save without one and have no persisted anchor.
`session.persistParagraphIds()` assigns them IDs across every story part, comments,
note separators and retained XML included, and repairs duplicates: source IDs and
saved IDs keep theirs over copies. It refuses, changing nothing, rather than guess at
an ambiguous comment reference. The change is replicated, stays out of undo history,
and later saves keep it; a save whose stories are otherwise unchanged patches the IDs
into the source bytes. `session.paragraphIdentities()` lists each paragraph's session,
persisted and exact-source anchors. A persisted anchor is scoped to the document the
host chose, and repeated source IDs resolve as `ambiguous`.

### Selection state

`session.selectionContext(range)` aggregates the range for toolbars and
assistive technology. Toggle marks, including `superscript` and `subscript`,
are `true`, `false` or `'mixed'`; value marks such as `fontFamily`, `color` and
`highlight` are `null` when absent or not uniform. `highlight` reports Word's
highlight name (`'yellow'`) or an unmapped hex value. `toggleMark` accepts
`{ type: 'superscript' }` and `{ type: 'subscript' }`; adding one clears the other.

### Version-checked edit batches

Read what you will target together with the session version, then apply a batch
against that version. Every step resolves against the state that was read; the
batch commits as one transaction or returns a typed refusal with nothing changed.

```ts
const read = session.readParagraphs({ story: 'body', view: 'accepted' });
if (!read.ok) throw new Error(read.failure.message);
const clause = read.paragraphs.find((paragraph) => paragraph.text.startsWith('Term'))!;

const result = session.applyEdits({
  expectVersion: read.version,
  steps: [
    {
      op: 'replaceText',
      target: { kind: 'search', text: '30 days', view: 'accepted',
        within: { kind: 'paragraph', story: 'body', paraId: clause.paraId } },
      text: '45 days',
      expect: { text: '30 days' },
    },
    {
      op: 'insertParagraphs',
      target: { story: 'body', paraId: clause.paraId },
      at: 'end',
      paragraphs: [{ text: 'Renewal is automatic.' }],
    },
  ],
});
if (!result.ok) console.warn(result.failure.code, result.failure.message);
```

`version()` changes with every committed change, local or remote (including
undo and redo), and when the session reopens its document. Tokens are scoped to
one session; after a lost response, read again rather than retrying with a new
version.

Offsets are UTF-16 positions into a paragraph's projected text. Each inline atom
(hard break, image, content control, note reference, field, other embed) is one
U+FFFC listed in `atoms`, tabs stay `\t`, and paragraph marks are excluded. The
`accepted` view includes pending insertions and hides pending deletions;
`original` does the reverse. `findText` is exact, case-sensitive and
paragraph-local, and a `search` target must match exactly once in its scope.

| Step | Effect |
| --- | --- |
| `insertText` | Inserts at the start or end of a target, formatted like typing there. |
| `replaceText` | Replaces a target's text inside one paragraph, keeping the paragraph. |
| `deleteText` | Deletes a target's text inside one paragraph. |
| `insertParagraphs` | Inserts complete paragraphs before or after an anchor, in the given or the anchor's style. |
| `deleteParagraphs` | Deletes a contiguous span of complete paragraphs without merging properties into a neighbour. |
| `setParagraphStyle` | Applies a document paragraph style with its paragraph and run formatting. |

Refusals carry a `code` (`stale-version`, `missing-target`, `ambiguous-target`,
`content-mismatch`, `overlapping-steps`, `locked-target`,
`tracked-revision-conflict`, `unsupported`, `invalid-step`, `limit-exceeded`),
the failing `stepIndex`, and the target. Editor hosts add `read-only` for an
editor that does not accept edits. Malformed requests throw.
`validateEdits(request)` runs the same checks and previews each step without
changing anything or reserving ids.

An applied batch is exactly one undo step in both capture modes; pass
`history: 'none'` to keep it out of undo history while preserving existing undo
and redo entries. `source: 'agent'` records provenance only. Add
`suggest: { author, date }` to a text step to record it as a tracked change.

Paragraph ids are session keys; inserted paragraphs get Word paragraph IDs as
typed ones do, so `saveYrsDocx` returns persisted anchors for them.

### Structured export

Export a document as read-only structured JSON or Markdown, with the location of
every block and inline and a diagnostic for everything the export omits or cannot
represent.

```ts
import { exportDocxMarkdown, exportDocxStructured } from '@betteroffice/docx';

const content = await exportDocxStructured(bytes, {
  revisionView: 'accepted',
  stories: ['body', 'headers', 'footers'],
});
for (const block of content.stories[0].blocks) {
  if (block.kind === 'heading') console.log(block.heading.outlineLevel, block.anchor);
}
const { markdown, anchors } = await exportDocxMarkdown(bytes, { revisionView: 'markup' });
```

On a live session, `session.exportStructured(options)` and
`session.exportMarkdown(options)` return `{ ok: true, version, content }` from one
read of the committed state: anchors resolve against that version, and nothing is
committed, flushed or published. Bytes exports and `renderDocxMarkdown(content)`
return snapshot content and throw `DocxExportError` for unusable options. A
refusal's `failure` is `{ code, target, message }`, with `target: null` when it
concerns no anchor; `unsupported` means the session holds no document content.
`session.headings(story)` lists a story's headings classified the same way; the
legacy `collectHeadings` tree walker from `@betteroffice/docx/utils` is deprecated
in its favour.

Content is `schemaVersion: 1`: ordered stories (`body` by default; `headers`,
`footers`, `footnotes`, `endnotes` and `comments` only when selected), each a list
of paragraphs, headings (outline level and whether it came from direct formatting,
the style chain, document defaults or a `HeadingN` style id), list items (numbering
format and the rendered marker, counted as Word counts them: numbering instances of
one abstract definition continue each other unless one overrides a start and so
begins its own list, levels begin at `w:start` and restart as `w:lvlRestart` says),
tables on the source grid with spans, skipped grid columns and vertical-merge
continuations, content controls, section breaks, and placeholders naming the
source element of other content. Inlines cover text, tabs, line,
page and column breaks where the source has them, note and comment references,
fields with their cached result in result order, hyperlinks and nested fields
included, all anchored to the field (never evaluated; a numeric field's result
blocks stay inside the field), images (alt text and the relationship of the part
that owns them, no binary data) and inline controls. A header or footer part
referenced through several relationships is one story with every section that
uses it while the copies hold the same content, and one story per copy with a
diagnostic once they differ. `revisionView` is required: `accepted` and
`original` project pending revisions like `readParagraphs`, and `markup` keeps
both with `revisions` attribution on each inline, moves as `moveFrom` and
`moveTo`. Where a view cannot reconstruct history (a paragraph-mark revision,
tracked formatting or a style change that changes the exported marks, tracked
cell or grid changes,
or row and table revisions the markup view cannot attribute) the block is an
anchored `unsupported` placeholder with an `unsupported-revision` diagnostic.

Anchors use the batch offsets and session keys: a paragraph anchor's `paraId` is the
batch target key, and a range is paragraph-local UTF-16 in the view it
names, one U+FFFC per atom, the same shape the edit batches take; ranges from a
session export address the version it returned. Content without a session
location (comment bodies, omitted raw XML) carries a `sourcePart` anchor: the
part, its SHA-256, and element-child ordinals into its XML. Content with no
location of its own carries `{ kind: 'unlocated', story, reason }`: the reason is
`duplicate-paragraph-id` (a paragraph whose Word paragraph ID repeats an earlier one
opens with a session key of its own, and edits and merges repair a key two paragraphs
share, so only shared state written outside the engine has it), `story-too-large`,
`missing-story` or
`provenance-unavailable`. Comment metadata reads
the session's comment store once a field has been written there, and the source
until then. Ids are deterministic export-tree paths, and a bytes export is
identical across runs. Exports admit one top-level block at a time, charge text before copying it, and stop at
`maxBlocks` (10,000 by default, every nested block counted) or `maxBytes`
(8,388,608 by default, on the compact JSON) at a whole top-level block, set
`truncated`, and end with a `truncated` diagnostic; the prefix already admitted is
kept, and a larger `maxBytes` never keeps fewer blocks. Of a story longer than
four units per remaining byte beyond one million, only the leading paragraphs
within that bound are read: those that fit are kept, then the export stops and says
so. A comment body, field payload, shape or table cell past the bound is not read
at all. Markdown keeps story separators, headings, lists, tables (entity-escaped
HTML, nested where needed, where Markdown tables cannot carry merges or several
blocks per cell), attributed insertions and deletions, and a
`<!-- docx-export:N -->` marker per block, nested ones included, mapped to its
anchor in `anchors`. Only http, https, mailto and internal-anchor targets are linked,
judged after entity and percent decoding, and `&` in a target is written `&amp;`; document text, titles and alt text never keep a
line break or unescaped HTML. It does not
preserve Word pagination or layout.

### Page fragments

A paged export attaches a page map to the structured content: the physical page,
section and displayed page label showing each block and inline, and the body,
header, footer or note occurrence it sits in. Ordinary exports never lay anything
out.

```ts
import { exportDocxStructuredWithPages, renderDocxMarkdownWithPages } from '@betteroffice/docx';

const paged = await exportDocxStructuredWithPages(
  bytes,
  { revisionView: 'markup', stories: ['body', 'headers', 'footnotes'], includeGeometry: true },
  { fonts: [{ key: 'sans', data: fontBase64 }], defaultChain: ['sans'] }
);
paged.layout.pages; // pageIndex, displayedLabel, sectionIndex, parityFiller, size
paged.layout.fragments; // pageIndex, occurrenceId, nodeId, slice, continuation flags, geometry
const { markdown } = await renderDocxMarkdownWithPages(paged, { pageMarkers: true });
```

`exportDocxStructuredWithPages` opens the bytes in a private session, registers
exactly the fonts it is given, as base64 data, in a measurement font store of its
own (a requirement `family|bold|italic` takes the chain named by its key, else by
its lowercase family, else `defaultChain`; nothing is looked up on the system and
no editor's fonts change), lays out every section, header, footer and note, and
exports from that one state. Its options are plain JSON: `measurementDefaults`
(`fontFamily`, `fontSize`), `renderEnvironment` (`showHiddenText`,
`defaultTabStopTwips`) and `compatibility`, all fingerprinted in the map's
provenance. A requirement left with no font, including the `measurementDefaults`
family text naming none falls back to, or any text measured with stand-in metrics
is refused as `layout-unavailable`, as is a live layout whose font store was
cleared or added to since.
The snapshot map is deterministic for the same bytes, fonts and options and
carries no session token. `session.exportStructuredWithPages(options)` reads the
region layout a session retains and lays nothing out, flushes nothing and loads no
font; the layout must have lowered the current version from the session's own
stories, with section, settings and note metadata that describe it and a
registered font for every requirement. Its map carries `documentVersion` and a
`layoutVersion` to pass back as `expectLayoutVersion`; its references describe the
authoritative layout of that version, which a later edit may supersede before it is
painted.

`pageIndex` counts physical pages from zero, blank parity pages included, while
`displayedNumber`/`displayedLabel` follow the section's PAGE numbering (restarts,
continuation and Roman or letter formats; an unwritable format falls back to
decimal with `numberingStatus: 'fallback'` and an `unsupported-numbering`
diagnostic). A header or footer part stays one exported story with an occurrence
on every page showing it; a note has an occurrence where its note area is, which
may differ from its reference's page. Text slices are ranges in the export's own
offsets, split where the node, paragraph, view or page changes; atoms (fields,
controls, images, note marks, breaks) are sliced whole and marked `partial` with
an `anchor-only` diagnostic when only part of their content is on the page. Table
fragments list their row window with `continuedFromPrevious`, `continuedOnNext`
and `repeatedHeader`, and each cell paragraph has fragments of its own for the lines
its cell shows (a line the cell cuts through is listed with a `clipped-content`
diagnostic). Pages are
laid out with revision markup, so while a laid-out story holds pending revisions a
paged export reads the `markup` view, and `accepted` or `original` return
`unsupported-revision-layout`. The other refusal codes are `stale-document`,
`stale-layout` (stale section, settings or note metadata, fonts, options or
`expectLayoutVersion`), `layout-unavailable`, `layout-not-converged` and
`unsupported`; a note taller than its page's note area carries an
`unsupported-note-layout` diagnostic.
Geometry is off by default; rectangles are unzoomed CSS pixels (96 per inch) from
the physical page's top-left corner. `maxFragments` (100,000 by default) and
`maxLayoutBytes` bound the map separately and mark it `truncated`; mapping stops a
page past the limit, so a truncated map leaves out the diagnostics of later pages
and of paragraphs no page shows. With
`pageMarkers`, Markdown follows each block marker with
`<!-- docx-pages: 0=i 1=ii -->` (labels percent-encoded) and writes no page break
into the text; a map from other content is refused. The map is `schemaVersion: 1`,
versioned apart from the content: additive fields keep the version, and a change
existing readers would misread bumps it.

### Content controls

List a document's content controls and fill its plain- and rich-text controls in
one version-checked batch.

```ts
const read = session.listContentControls();
if (!read.ok) throw new Error(read.failure.message);
const name = read.content.controls.find((control) => control.tag === 'customer.name')!;

const result = session.applyEdits({
  expectVersion: read.version,
  steps: [
    { op: 'setContentControlText', target: { kind: 'id', controlId: name.controlId }, text: 'Ada Lovelace' },
    { op: 'setContentControlText', target: { kind: 'tag', tag: 'customer.address' }, text: '12 Example Street\nLondon' },
  ],
});
```

Controls come in document order (body, headers, footers, footnotes, endnotes,
comments; a control before the controls inside it), each with the structured
export's control metadata (`controlId`, `ooxmlId`, `controlType`, `tag`, `alias`,
`lock`, `showingPlaceholder`, `dataBound`) plus `placement`, `anchor`,
`parentControlId`, `value` (its text with tabs as `\t` and line and paragraph
breaks as LF, or `unavailable` with a reason), `multiLine` and `effectiveLock`,
which folds in the locks of the controls containing it. `stories` narrows the read
to some categories; `maxControls` (10,000) and `maxBytes` (8,388,608) refuse with
`limit-exceeded` rather than returning part of the list, and `complete: false`
comes with diagnostics naming what could not be covered. Tags, aliases and
`ooxmlId`s match exactly and case-sensitively, and `findContentControls` returns
every match. `listDocxContentControls(bytes)` and
`findDocxContentControls(bytes, query)` read bytes without a session and throw
`DocxContentControlsError` for refused options. Concurrent fills from
collaborators resolve last-writer-wins for an inline control and merge like
concurrent typing for a block control.

A step's `target` is `{ kind: 'id', controlId }`, `{ kind: 'tag', tag }` or
`{ kind: 'ooxmlId', ooxmlId }`. `controlId` is the per-version locator: scoped to
the version or snapshot it was read at, it does not survive save and reopen, and
an inline control's id can name a different control after edits, so list again
after a version change. Tags are the template author's names for
controls. `ooxmlId` is the control's authored `w:id`, the identity that survives
save and reopen. A tag or `ooxmlId` write searches every story regardless of read
filters and needs exactly one control in the document to carry it.

`setContentControlText` replaces the content of a plain- or rich-text control
with plain text and clears its placeholder flag everywhere it is recorded,
leaving tags, aliases, ids, bindings and every other property as captured. CRLF
becomes LF; LF is a line break, except in a rich-text block control, where it
starts a paragraph: surviving paragraphs keep their ids and properties and new
ones take the first paragraph's style defaults. A plain-text control accepts LF
only when its `w:text` sets `w:multiLine` and keeps its lines in one paragraph.
The text takes the formatting of the control's first text run, or the control's
own run properties while it shows its placeholder. Equal text is a no-op unless
the placeholder still shows or a plain-text control holds several paragraphs,
which the fill joins into one. Refusals
carry the batch code and a `reason`: `missing-control`, `missing-tag` or
`missing-ooxml-id`, `ambiguous-control-id`, `ambiguous-tag` or
`ambiguous-ooxml-id`, `content-locked`, `bound-control`,
`unsupported-control-type`, `unsupported-children` (content other than text,
tabs and line breaks, such as bookmarks, fields or tables), `nested-controls`,
`unknown-lock`, `unsupported-suggestion`, `provenance-unavailable` (the session
was not opened from DOCX bytes), `unsupported-story` (a control kept only in
source XML, such as a comment body), `multiline-not-allowed` and `invalid-text`.
Checkbox, dropdown and date controls keep `setContentControlValue`; a
string passed to it fills a text control through the same step, and
`clearContentControlValue` never erases a text control's text.

### Compare documents into tracked changes

Compare an original and a revised DOCX and get back the original with the
body text differences as tracked insertions and deletions, attributed to the
author and date you pass:

```ts
import { compareDocx } from '@betteroffice/docx';

const result = await compareDocx(original, revised, {
  author: 'Contract review',
  date: '2024-05-06T09:30:00+02:00',
  granularity: 'word',
  unsupported: 'fail',
});
if (result.ok) {
  for (const change of result.changes) console.log(change.id, change.kind, change.revised.text);
} else {
  console.warn(result.diagnostics.map((diagnostic) => diagnostic.code));
}
```

`compareDocx` compares the text of body paragraphs whose structure is unchanged. Both
packages are inspected in full before anything is authored, and any blocking
difference refuses the comparison with `{ ok: false, diagnostics }`; no partial
redline is ever returned. `unsupported: 'report'` keeps inspecting and returns
every diagnostic, bounded by `limits.maxDiagnostics`. Blocking codes are
`invalid-options`, `invalid-docx`, `existing-revisions` (anywhere, headers and
property revisions included), `paragraph-insertion`, `paragraph-deletion`,
`paragraph-move`, `ambiguous-alignment`, `table-change`, `structure-change`,
`field-change`, `object-change`, `content-control-change`, `formatting-change`,
`unsupported-content` and `unsupported-formatting` (a changed paragraph holding
content, or run properties, that edits cannot carry, such as fields,
hyperlinks, bookmarks, objects, page breaks, patterned shading or `w:lang`),
`out-of-scope-change` (headers, footers, notes, comments), `opaque-part-change`,
`provenance-unavailable`, `limit-exceeded`, `diagnostics-truncated`,
`batch-refused`, `serialization-failed` and `roundtrip-mismatch`.
`metadata-difference` (informational: modified dates, revision-session ids and
document statistics, kept from the original) and `ambiguous-identity` (a
warning: repeated paragraph ids, so changes are addressed by position) accompany
a result. Every part whose content type or content is XML is scanned for
revisions, and wrappers such as custom XML elements, body-level markers and
every part must match exactly; an XML-typed part that cannot be read refuses the
comparison.

Paragraphs are aligned conservatively: text unique to one paragraph on each side
anchors the alignment, crossing anchors are moves, one paragraph against one
between anchors is a replacement, and several need shared words and exactly one
best correspondence. Clearing a paragraph's text is a deletion; the paragraph
mark stays. Differences are reported by Unicode word, keeping punctuation and
whitespace tokens, or with `granularity: 'char'` by grapheme cluster. Each
change is one tracked replacement: deleted text keeps its formatting and
inserted text takes the revised formatting. Changes are ordered by original
position, `id` is `change-0`, `change-1` and so on, and their spans address the
inputs by part, SHA-256 and element-child path with paragraph-local UTF-16
offsets. `date` must be RFC 3339 with an explicit offset and is recorded in UTC;
the clock is never read. `author` must not be blank.

Only the changed paragraphs' content is rewritten; their start tags and
properties, every other paragraph and every other part keep the original's
bytes, and identical inputs return the original bytes. Rewritten runs carry the
session's formatting vocabulary, so run properties a style supplies are written
as direct formatting, and complex-script fonts, sizes, bold and italic are
written only as the source states them. The result is reopened and checked
against both inputs before it is returned, with each run's effective formatting
read from the XML, style toggles such as bold combined across every level of a
style's `basedOn` chain; a difference refuses with `roundtrip-mismatch`.
Defaults, which are also the ceilings: 32 MiB per input, 128 MiB inflated,
10,000 paragraphs per input and 1,048,576 UTF-16 units of text in both inputs,
counted across every story while the parts are read, 250,000 alignment cells,
4,000,000 diff cells, 128 changes, 256 diagnostics, 64 MiB of staged state,
8 MiB of result JSON without the output package (at least 1 KiB) and a 64 MiB
output, the no-op included.
Review in BetterOffice is tested.

### Editor plugin contract (deprecated)

`@betteroffice/docx/plugin-api` keeps its geometry and sidebar building blocks:
`RenderedDomContext`, `PositionCoordinates`, `SidebarItem`,
`createRenderedDomContext`, `createCanvasHostProjector` and
`resolveItemPositions`. The snapshot-based `EditorPluginCore`,
`PluginPanelProps`, `PanelConfig` and `SidebarItemContext` are deprecated: they
pass serialized `Document` snapshots and have no host that manages their
lifecycle. Host plugins for the React editor use `defineDocxPlugin` and the
`plugins` prop of
[`@betteroffice/docx-react`](https://www.npmjs.com/package/@betteroffice/docx-react),
which provide versioned reads, granted commands and edit batches, lifecycle
events and cleanup.

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

`saveYrsDocx()` writes the current ranges under the comment's own ID, with its
author, date, body, replies and resolution. Reopening the saved file restores a
range that stays within one paragraph and overlaps no other comment's range;
opening reads comment ranges paragraph by paragraph, one comment per run.

Reanchoring joins the current local undo capture, so an immediate replacement and
reanchor can undo together. Comment edits participate in the session's history;
undoing comment changes conservatively invalidates all stories for saved anchors.

The undo manager retains comment item boundaries with their undo/redo entries so
anchors inside replaced text survive repeated history traversal with Yrs 0.27.
Discarding history releases its bookkeeping; undo/redo traverses a local snapshot
when those boundaries need restoration.

### External drop positions

`RenderedDomContext.getPositionAtPoint(clientX, clientY)` queries the same canvas
hit testing used for the caret. Pass viewport coordinates, such as a drop event's
`clientX` and `clientY`. Page scroll and CSS zoom are resolved from the live page
geometry. The query never changes selection or focus.

The result contains `position`, `pageIndex`, and `region`. Header/footer hits also
carry `rId`; footnote/endnote hits carry `noteId`. Positions are local to that
region's display stream in the layout the context renders, and shift with every
edit.

In React, `editorRef.getPositionAtPoint(clientX, clientY)` runs the same query and
adds `version` and `target`: the collapsed accepted-view range at the point, keyed
by session paragraph keys, in the body, header, footer, note or table cell story
the point is in. Use it as a text step's target with `expectVersion: version`:

```ts
const hit = editorRef.current!.getPositionAtPoint(event.clientX, event.clientY);
if (hit) {
  await editorRef.current!.applyEdits({
    expectVersion: hit.version,
    steps: [{ op: 'insertText', target: hit.target, at: 'start', text: '{{customer.name}}' }],
  });
}
```

It returns `null` while typed or composed input is pending and until the painted
pages show the current version (retry after `flushPendingInput()` or on the next
frame), and a batch applied after further typing refuses with `stale-version`, so a
drop never lands at a shifted position. Plugins get the same query as
`context.geometry.getPositionAtPoint`, with the layout's `layoutId`. The lower-level
`editorRef.getEditorRef()?.displayPositionToYrsLoc(hit)` maps a hit to its live
`YrsLoc` (offsets that count text a pending deletion hides) without the version
check; passing a number to it retains body mapping.

Text runs and their editable content boxes are accepted. Page margins, images,
page gaps, points outside pages, and queries without ready canvas geometry return
`null`.
