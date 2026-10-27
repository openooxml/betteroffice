# @betteroffice/xlsx-react

React XLSX editor with a canvas grid, toolbar, selection, keyboard input, and clipboard support.
The engine runs in a Web Worker by default, with promise-returning host operations and exports.

## Install

```bash
bun add @betteroffice/xlsx-react react react-dom
```

`react` and `react-dom` 18 or 19 are peer dependencies; the XLSX core is included as a dependency.

## Usage

```tsx
import { XlsxEditor } from '@betteroffice/xlsx-react';

export function Spreadsheet({ bytes }: { bytes: Uint8Array }) {
  return (
    <div style={{ height: 600 }}>
      <XlsxEditor file={bytes} fileName="report.xlsx" />
    </div>
  );
}
```

`file` contains XLSX bytes. Give the editor's container a height; omitting `file` shows an empty
editor that asks for an XLSX file. See [hosting](https://docs.betteroffice.dev/docs/migrating-to-0-5#hosting) for worker, chunk and CSP needs.

## Key options

The default props are described by `XlsxAsyncEditorProps`:

| prop | type | description |
| --- | --- | --- |
| `file` | `Uint8Array \| ArrayBuffer` | XLSX bytes to open. |
| `fileName` | `string` | Download name; defaults to `workbook.xlsx`. |
| `onReady` | `(api: XlsxAsyncEditorApi) => void \| (() => void)` | Receive the editor API; optionally return cleanup. |
| `onSave` | `(bytes: Uint8Array) => void` | Receive saved workbook bytes. |
| `onSaveRequest` | `() => boolean \| void \| Promise<boolean \| void>` | Intercept saves; return `true` to continue. |
| `downloadOnSave` | `boolean` | Save also downloads; defaults to `!onSave`. |
| `onDocumentChange` | `(change: XlsxDocumentChange) => void` | Receives `{ version }` after a committed edit. |
| `initialZoom` | `number` | Initial scale, clamped to 0.1 through 5; defaults to 1. |
| `loadingIndicator` | `React.ReactNode` | Shown while a supplied file opens. |
| `style` | `React.CSSProperties` | Overrides root inline styles. |
| `onChange` | `() => void` | Called after a user edit changes the workbook. |
| `onError` | `(error: Error) => void` | Receive engine errors. |
| `readOnly` | `boolean` | Block user edits while keeping navigation and selection. |
| `toolbar` | `React.ReactNode` | Replace the default toolbar; `null` removes it. |
| `showToolbar` | `boolean` | `false` hides the toolbar region, including custom content. |
| `collaboration` | `XlsxEditorCollaborationOptions` | Configure the editor-owned replica and presence provider. |
| `i18n` | `Translations` | Override UI strings using `@betteroffice/xlsx-i18n`. |
| `plugins` | `readonly XlsxPlugin[]` | Install host plugins. |
| `pluginGrants` | `Readonly<Record<string, XlsxPluginGrant>>` | Grant capabilities by plugin ID. |
| `onPluginError` | `(error: XlsxPluginError) => void` | Receive plugin failures. |
| `className` | `string` | CSS class for the editor root. |

All listed props are optional. The default toolbar is hidden when `readOnly`; a supplied `toolbar`
remains visible unless `showToolbar` is `false`.

Deprecated: `legacySyncApi` keeps the synchronous API on the in-thread engine and is removed in 0.6.0; omit it
for the default asynchronous editor API.

## Editor API

`onReady` receives `XlsxAsyncEditorApi`. Its exports wait for pending editor input before reading the
workbook, and its command store is shared with the toolbar and host controls.

| member | purpose |
| --- | --- |
| `hydrated` | Whether the worker editor is ready. |
| `whenHydrated()` | Await editor readiness. |
| `failure` | Current engine failure, or `null`. |
| `commands` | Inspect command availability and execute actions asynchronously. |
| `version()`, `readCells`, `findText` | Read after pending input commits. |
| `validateEdits`, `applyEdits` | Validate or apply version-checked batches; refuse while read-only. |
| `save()` | Commit input, return bytes and call `onSave` without downloading or consulting `onSaveRequest`. |
| `focus()` | Focus the editor. |
| `renderPng`, `renderRangePng`, `printDisplayList` | Render asynchronously. |
| `exportStructured`, `exportMarkdown` | Export content asynchronously. |
| `proposeChanges` | Optional method for staging tracked cell edits on worker editors. |
| `collaboration` | Asynchronous collaboration replica, or `null`. |
| `recoverySave()` | Recover bytes after a non-collaborative worker session fails. |

Treat the API as belonging to the open document. Return cleanup from `onReady` to release host
subscriptions when that document is replaced or the editor unmounts.

## Saving

Use `await api.commands.execute('save', null)` to save from host UI. The command waits for accepted
input and calls `onSave` when supplied. It also downloads when `downloadOnSave ?? !onSave` is true.
Use `await api.save()` to receive bytes directly without downloading or consulting `onSaveRequest`.

`onSaveRequest` intercepts toolbar saves, the save command, and Ctrl/Cmd+S before serialization.
Return `true` to continue built-in saving; `false` or `void` handles or cancels the request. Promises
are awaited, concurrent requests are coalesced, and requests for replaced documents are discarded.

Commands return a success or a typed failure. A refused cell draft must be corrected or discarded
before saving can complete. `recoverySave()` requires a failed non-collaborative worker session and
returns `{ bytes, recovery: true }`.

## Toolbar

The toolbar composition API is experimental. Use `toolbar` to supply controls connected to the
editor's command store. See
[Compose the XLSX toolbar](https://docs.betteroffice.dev/docs/javascript#compose-the-xlsx-toolbar)
for built-in parts, custom controls, and command results.

## Plugins and edit batches

The plugin API is experimental and may change in minor releases. Plugins contribute panels,
overlays, and commands through restricted clients; their read and edit
clients support version-checked operations. See
[Host XLSX editor plugins](https://docs.betteroffice.dev/docs/javascript#host-xlsx-editor-plugins)
and [Edit XLSX in version-checked batches](https://docs.betteroffice.dev/docs/javascript#edit-xlsx-in-version-checked-batches).

## Structured export

Await `api.exportStructured()` or `api.exportMarkdown()` for content with positional anchors and
omission diagnostics. See
[Export XLSX content with anchors](https://docs.betteroffice.dev/docs/javascript#export-xlsx-content-with-anchors)
for scopes, hidden content, and export limits.

## Collaboration

Pass `collaboration` and connect a provider when `collaboration.onReplica` publishes the replica.
The editor retains its default worker runtime; see the
[collaboration guide](https://docs.betteroffice.dev/docs/collaboration) for transport setup and cleanup.

## Documentation

- [JavaScript guide](https://docs.betteroffice.dev/docs/javascript)
- [Collaboration guide](https://docs.betteroffice.dev/docs/collaboration)
- [Package guide](https://docs.betteroffice.dev/docs/packages)
- [XLSX core on npm](https://www.npmjs.com/package/@betteroffice/xlsx)
- [Localization package on npm](https://www.npmjs.com/package/@betteroffice/xlsx-i18n)
- [Repository](https://github.com/openooxml/betteroffice)
