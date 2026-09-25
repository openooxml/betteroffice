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

Current limits: text targets stay within one paragraph; inline atoms cannot be
replaced or deleted; content-locked controls, existing tracked changes a step
would touch (including tracked run formatting), and pending paragraph-mark
revisions refuse. Paragraph and style steps need a session opened from DOCX
bytes (`openDocx`/`seedFromDocx`) and cannot be suggested. List numbering is a
v1 limitation: restyling a numbered paragraph, or applying or inserting a style
that defines numbering, refuses with `unsupported` until a follow-up retains
numbering definitions in the session. Deleting spans that hold tables,
controls, section breaks, opaque XML, fields' cached results, or comments and
bookmarks crossing the span refuses, as does any step after which saving would
move an opaque XML block, such as one that precedes a table. A batch holds at most 128 steps,
1,048,576 inserted UTF-16 units and 1,024 new paragraphs. Paragraph ids are
session anchors: they are not guaranteed to survive save and reopen.

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
begins its own list, levels begin at `w:start` and restart as `w:lvlRestart` says,
and a number a format cannot write, such as a Roman numeral past 3,999, is left
unresolved with a diagnostic), tables on the source grid with spans, skipped grid columns
and vertical-merge continuations, content controls, section breaks, and
placeholders for content v1 does not represent. Inlines cover text, tabs, line,
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

Anchors use the batch offsets: a range is paragraph-local UTF-16 in the view it
names, one U+FFFC per atom, the same shape the edit batches take; ranges from a
session export address the version it returned. Content without a session
location (comment bodies, omitted raw XML) carries a `sourcePart` anchor: the
part, its SHA-256, and element-child ordinals into its XML. Parsing gives a
repeated paragraph id a fresh one; paragraphs a session gives a shared id are
anchored with an empty `paraId`, which addresses nothing. Comment metadata reads
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
preserve Word pagination or layout. Page fragments are not included yet.

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

Only text inside body paragraphs whose structure is unchanged is compared. Both
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
a result. Parts are found by their content types, and wrappers such as custom
XML elements, body-level markers and every part must match exactly; a part that
cannot be read refuses the comparison.

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
8 MiB of result data and a 64 MiB output, the no-op included.
Review in BetterOffice is tested; Word validation is reported separately.
Native Rust and Python comparison is not available yet.
