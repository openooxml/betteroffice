# @betteroffice/docx-react

React DOCX editor and viewer with a toolbar, paginated canvas, comments, and tracked changes.
The engine runs in a Web Worker by default; document reads, edits, and exports return promises.

## Install

```bash
bun add @betteroffice/docx-react react react-dom
```

Peer dependencies: `react` and `react-dom` 18 or 19.

## Usage

```tsx
import { DocxEditor } from '@betteroffice/docx-react';
import '@betteroffice/docx-react/styles.css';

export function Editor({ file }: { file: File }) {
  return <DocxEditor documentBuffer={file} />;
}
```

Import the stylesheet once at the application entry, page, or layout level. Use a browser-only
component when integrating with server rendering.

## Key options

All props below are optional. Omit the document source to show the placeholder.

| prop | type | description |
| --- | --- | --- |
| `documentBuffer` | `ArrayBuffer \| Uint8Array \| Blob \| File \| null` | DOCX source bytes. |
| `document` | `Document \| null` | Parsed model from `@betteroffice/docx`, as an alternative source. |
| `mode` | `'editing' \| 'suggesting' \| 'viewing'` | Direct edits, tracked changes, or viewing; defaults to `'editing'`. |
| `readOnly` | `boolean` | Disable editing and hide the default toolbar and rulers. |
| `author` | `string` | Author for comments and tracked changes. |
| `onSave` | `(buffer: ArrayBuffer) => void` | Receive exported DOCX bytes. |
| `onSaveRequest` | `() => boolean \| void \| Promise<boolean \| void>` | Own the UI save workflow before serialization. |
| `downloadOnSave` | `boolean` | Also download UI saves; defaults to `true`. |
| `onDocumentChange` | `(change: DocxDocumentChange) => void` | Receive the document version after a committed change. |
| `onError` | `(error: Error, info?: DocxErrorInfo) => void` | Receive errors; recoverable reports include `recoverable: true`. |
| `collaboration` | `DocxEditorCollaborationOptions` | Configure the shared replica. |
| `showToolbar` | `boolean` | Show built-in or supplied chrome; defaults to `true`. |
| `toolbar` | `ReactNode` | Replace the toolbar; `null` hides it. Experimental composition API. |
| `showRuler` | `boolean` | Show the ruler; defaults to `false`. |
| `showZoomControl` | `boolean` | Show zoom controls; defaults to `true`. |
| `showHiddenText` | `boolean` | Include hidden text in layout; defaults to `false`. |
| `previewFirstPage` | `boolean` | Experimental early page preview; off by default and ignored with collaboration. |
| `onFirstPagePainted` | `() => void` | Called once per document when its first pages paint. |
| `pageMirror` | `'full' \| 'text'` | Accessible page DOM; defaults to `'full'`. |
| `measurementFontProvider` | `BundledFontProvider` | Override the measurement font provider for this editor. |
| `i18n` | `Translations` | Locale overrides from `@betteroffice/docx-i18n`. |
| `plugins` | `readonly DocxPlugin[]` | Host extensions with managed lifecycle and permissions. |
| `memoryBudget` | `DocxMemoryBudget` | Memory pressure thresholds and an optional worker allocation limit. |
| `onMemoryPressure` | `(pressure: DocxMemoryPressure) => void` | Receive memory pressure level changes. |

Deprecated: `legacySyncApi` opts into the synchronous hybrid runtime and is removed in 0.6.0;
omit it to use the default async editor.

Deprecated: `onChange`; use `onDocumentChange` for version notifications. `experimentalWorkerOpen` is
deprecated too; the worker editor is the default.

## Editor ref

`DocxEditorRef` exposes async document operations and the editor's command store, `commands`.
Await `save()` to export DOCX bytes, or `flushPendingInput()` before a separate host operation
that must include accepted typing and the current selection. Flushing waits for active IME
composition and rejects if input fails, the document is replaced, or the editor unmounts.

Use `readPositionAtPoint(clientX, clientY)` for a versioned text target at viewport coordinates.
Use `scrollToParagraph(paraId)` for async paragraph navigation. Both support the default worker
runtime.

Deprecated: `getPositionAtPoint` and `scrollToParaId`; use `readPositionAtPoint` and
`scrollToParagraph`, respectively.

## Host saves

`onSaveRequest` handles File > Save and Cmd/Ctrl+S before serialization. Return `true` to continue
the built-in save; return `false` or nothing to suppress it. Async callbacks are awaited, overlapping
UI requests share one workflow, and errors reach `onError`.

Calling `save()` inside that callback exports directly without invoking `onSaveRequest` again.
`onSave` receives the exported buffer; `downloadOnSave={false}` disables the additional UI download.

## Loading and accessibility

`previewFirstPage` paints a display-only preview while the document opens. Editing and plugin
activation wait for the full document. `onFirstPagePainted` lets a host dismiss its loading UI
when either preview or full-document pages first appear.

`pageMirror="text"` keeps readable text, links, and tables with less DOM work. It omits on-page
text positions for assistive technology and elements for plugin DOM range queries.

`getTotalPages()` returns the current, possibly provisional, page count. Await
`whenLayoutComplete()` for the final count; pages outside the viewport are built when shown.

## Host search

`search(query, options?)` on the ref highlights body matches, including tables, without moving
selection or focus. It returns a promise and works in viewing mode; matching is case-insensitive
unless `caseSensitive` is set. An empty query clears the search.

`searchNext()` and `searchPrevious()` wrap through matches; `searchGoTo(index)` selects a match.
Read state with `getSearchState()`, subscribe with `onSearchChange()`, and clear with `clearSearch()`.

## Host proposals

The ref's `proposeChanges()`, `setProposalStates()`, `withdrawProposals()`, and `getProposals()`
manage tracked changes grouped by host proposal IDs. Decisions preview accepted proposals as plain
text and hide rejected ones; withdrawal settles those decisions outside undo history.

Set `allowHostProposals` to permit proposals in read-only or viewing mode. Set
`showHostProposalsInSidebar` as well to display proposal cards. Both are off by default;
proposal access does not enable typing, edit batches, commands, or plugin writes.

## Toolbar and plugins

Use the experimental composition API to arrange built-in controls or supply external chrome;
see [compose the DOCX toolbar](https://docs.betteroffice.dev/docs/javascript#compose-the-docx-toolbar).
The plugin API is experimental too. For host panels, overlays, sidebar items, lifecycle, and
granted operations, see
[host DOCX editor plugins](https://docs.betteroffice.dev/docs/javascript#host-docx-editor-plugins).

## Edit batches and paragraph anchors

The ref's `readParagraphs`, `findText`, `validateEdits`, and `applyEdits` flush pending input before
version-checked operations. See [DOCX edit batches](https://docs.betteroffice.dev/docs/javascript#edit-docx-in-version-checked-batches)
and [paragraph anchors](https://docs.betteroffice.dev/docs/javascript#docx-paragraph-anchors)
for target addressing, persisted identities, and refusals.

## Structured export and content controls

For `exportStructuredWithPages(options)` in the default worker editor, see
[page fragments](https://docs.betteroffice.dev/docs/javascript#page-fragments) for revision views,
layout requirements, and geometry.

The ref's `listContentControls` and `findContentControls` inspect controls before version-checked
fills; see [content controls](https://docs.betteroffice.dev/docs/javascript#fill-docx-content-controls).

## Collaboration

The `collaboration` prop connects the worker editor to a shared replica. See the
[collaboration guide](https://docs.betteroffice.dev/docs/collaboration) for replica callbacks,
shared initial state, transport wiring, and checkpoints.

## Fonts and localization

Configure optional measurement fonts before mounting editors; see
[font setup](https://docs.betteroffice.dev/docs/packages#fonts). Use `measurementFontProvider` for
an editor-specific provider. Embedded font faces are isolated between documents and released when
the editor unmounts or loads another document.

Pass locale strings through `i18n`; strings and translation types are published in
[`@betteroffice/docx-i18n`](https://www.npmjs.com/package/@betteroffice/docx-i18n).

## Memory and recovery

`getMemoryStats()` reports main-thread and worker WebAssembly memory. `onMemoryPressure` reports
threshold crossings; `memoryBudget` defaults to warning at 75% and critical at 90% of 4 GiB.
`memoryBudget.workerLimitBytes` caps worker editing-core allocations.

A failed worker can be replaced once per editor, reopening the source and replaying edits.
For byte sources, recovery uses the last checkpoint, so undo history reaches back to that checkpoint.
If recovery fails, `onError` reports the failure and editing stops. Collaboration sessions report
worker failure without automatic replacement.

## Documentation

- [JavaScript guide](https://docs.betteroffice.dev/docs/javascript)
- [Collaboration guide](https://docs.betteroffice.dev/docs/collaboration)
- [Package guide](https://docs.betteroffice.dev/docs/packages)
- [Core package on npm](https://www.npmjs.com/package/@betteroffice/docx)
- [Locale strings on npm](https://www.npmjs.com/package/@betteroffice/docx-i18n)
- [Repository](https://github.com/openooxml/betteroffice)
