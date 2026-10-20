# @betteroffice/docx

Framework-free DOCX parsing, serialization, CRDT editing, layout, and canvas rendering backed by
Rust/WebAssembly. Use it for document processing or a custom editor adapter.

## Install

```bash
bun add @betteroffice/docx
```

`@betteroffice/fonts` is an optional peer dependency for font assets.

## Usage

Parse a DOCX and save it while preserving untouched package parts:

```ts
import { readFile, writeFile } from 'node:fs/promises';
import { parseDocx, repackDocx } from '@betteroffice/docx';

const document = await parseDocx(await readFile('input.docx'));
// document.package.document.content: the body's paragraphs and tables
// document.package also holds styles, numbering, theme, media, headers and footers
const bytes = await repackDocx(document);
await writeFile('output.docx', new Uint8Array(bytes));
```

## Entry points

| entry point | description |
| --- | --- |
| `parseDocx` | Read an `ArrayBuffer`, `Uint8Array`, `Blob`, or `File` into a document model. |
| `repackDocx`, `repackDocxWithWarnings` | Save a parsed model against its original package; the latter also returns save warnings. |
| `createDocx` | Write a new document model without an original package. |
| `createYrsSession` (from `@betteroffice/docx/yrs`) | Asynchronously create a CRDT editing session. |
| `saveYrsDocx` (from `@betteroffice/docx/yrs`) | Save a session opened from DOCX bytes, returning bytes and persisted paragraph anchors. |

For a React editor and viewer, use
[`@betteroffice/docx-react`](https://www.npmjs.com/package/@betteroffice/docx-react).

## Runtime and assets

The package ships container, parser, layout, and editing WebAssembly assets in `dist/generated/`.
Async entry points load the assets they need; no manual initialization is required for parsing or
creating a session. Preserve the packaged asset URLs when bundling for a browser.

Core sessions expose synchronous editing and read methods after asynchronous creation. The React
editor runs its engine in a Web Worker by default and exposes promise-returning document reads,
edits, and exports. Release a direct session with `session.destroy()` when finished.

## Undo grouping

A story is one text flow: the body, a header or footer, a footnote, or a comment.
`session.setUndoCaptureMode(mode)` selects how local transactions group into undo steps;
`session.undoCaptureMode()` reads the mode. Call `session.beginUndoCapture()` after import or seeding.

| Mode | Grouping |
| --- | --- |
| `auto` (default) | Nearby edits coalesce; switching stories closes the group. |
| `manual` | Edits coalesce until an explicit boundary, including across stories. |

`session.addUndoBoundary()` closes the current group. Changing modes also closes it while keeping
history; setting the same mode does not. In manual mode, close groups before unrelated edits.
Remote transactions stay outside local undo history; edit batches supply atomic execution.

## Selection state

`session.selectionContext(range)` aggregates formatting across a paragraph-addressed selection.
Toggle formatting is `true`, `false`, or `'mixed'`; value formatting is `null` when absent or
nonuniform. Use this state to drive custom formatting controls.

## Host proposals

`session.proposeChanges({ expectVersion, proposals })` applies a version-checked round of tracked
changes grouped by host proposal IDs, outside undo history. Each proposal names its `id`, target
`paragraph`, an `op` (`replaceText` or `insertText`) and `suggest: { author, date }` for attribution.
A refusal leaves the round unapplied.

`session.setProposalStates({ expectVersion, expectPreviewVersion, changes })` changes how proposals
preview without changing the document or undo history. `session.withdrawProposals({ expectVersion,
ids })` settles accepted proposals as plain text and removes rejected or undecided ones, so later
rounds search the text shown by the preview.

Read the registry with `session.getProposals()` or observe it with `session.onProposalChange()`.
Opening another document clears the registry. Use `proposalRevisionPreview(snapshot)` from
`@betteroffice/docx/yrs` as the layout render environment's `revisionPreview` to display decisions.

## Comment ranges

`session.setCommentRanges(commentId, ranges)` replaces an existing comment's sticky anchors while
preserving its metadata. Ranges may span paragraphs or address separate stories; each must be
nonempty, ordered, and within existing paragraphs. Invalid input throws before changing the comment.

Reanchoring joins local undo capture. `saveYrsDocx(session)` writes the current ranges; a host that
removes all commented text must choose a surviving range or remove the comment.

## Edit batches

Read targets with their session version, then apply all steps atomically against that version.
See [version-checked DOCX batches](https://docs.betteroffice.dev/docs/javascript#edit-docx-in-version-checked-batches)
for operations, offsets, limits, and typed refusals.

## Paragraph anchors

Session paragraph keys identify paragraphs within a collaborative session; persisted anchors identify
saved paragraphs by package part and OOXML paragraph ID. See
[DOCX paragraph anchors](https://docs.betteroffice.dev/docs/javascript#docx-paragraph-anchors)
for saving, reopening, resolving, and assigning IDs.

## Structured export

For `exportDocxStructured` and `exportDocxMarkdown`, see
[structured content](https://docs.betteroffice.dev/docs/javascript#export-docx-as-structured-content)
for anchored JSON or Markdown, revision views, story selection, and live-session exports.

For `exportDocxStructuredWithPages`, see
[page fragments](https://docs.betteroffice.dev/docs/javascript#page-fragments) for font requirements,
layout maps, and page geometry.

Deprecated: `collectHeadings` from `@betteroffice/docx/utils`; use `session.headings(story)` instead.

## Content controls

`listDocxContentControls` and `findDocxContentControls` inspect controls in DOCX bytes.
See [content controls](https://docs.betteroffice.dev/docs/javascript#fill-docx-content-controls)
for session discovery, version-checked fills, and supported control types.

## Compare documents

`compareDocx` compares two DOCX files and returns tracked text changes or diagnostics.
See [DOCX comparison](https://docs.betteroffice.dev/docs/javascript#compare-docx-files-into-tracked-changes)
for supported inputs, attribution, and refusal behavior.

## Collaboration

`CollaborationProvider` from `@betteroffice/docx/collaboration` connects a replica to a binary
transport. See the [collaboration guide](https://docs.betteroffice.dev/docs/collaboration)
for shared state, transport wiring, checkpoints, and deployment.

## Editor plugins

For React host extensions, use the
[DOCX plugin API](https://docs.betteroffice.dev/docs/javascript#host-docx-editor-plugins).

Deprecated: `EditorPluginCore`, `PluginPanelProps`, `PanelConfig`, and `SidebarItemContext` from
`@betteroffice/docx/plugin-api`; use `defineDocxPlugin` and `plugins` from `@betteroffice/docx-react`.

## Fonts for custom adapters

See [font setup](https://docs.betteroffice.dev/docs/packages#fonts) for optional measurement assets.
For browser font loading, `createFontLoadScope()` from `@betteroffice/docx/utils` isolates each
editor's load notifications. `registerDocumentFaces(faces, scope)` registers embedded faces and
returns their CSS family aliases when documents use different bytes under the same family name.
Release the scope with `scope.dispose()` when its editor is removed.

## Documentation

- [JavaScript guide](https://docs.betteroffice.dev/docs/javascript)
- [Collaboration guide](https://docs.betteroffice.dev/docs/collaboration)
- [Package guide](https://docs.betteroffice.dev/docs/packages)
- [React editor on npm](https://www.npmjs.com/package/@betteroffice/docx-react)
- [Locale strings on npm](https://www.npmjs.com/package/@betteroffice/docx-i18n)
- [Repository](https://github.com/openooxml/betteroffice)
