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

`file` contains XLSX bytes. Give the editor's container a height; omitting `file` shows a placeholder
grid.

## Key options

The default props are described by `XlsxAsyncEditorProps`:

| prop | type | description |
| --- | --- | --- |
| `file` | `Uint8Array` | XLSX bytes to open. |
| `fileName` | `string` | Download name; defaults to `workbook.xlsx`. |
| `onReady` | `(api: XlsxAsyncEditorApi) => void \| (() => void)` | Receive the editor API; optionally return cleanup. |
| `onSave` | `(bytes: Uint8Array) => void` | Receive saved bytes instead of downloading them. |
| `onSaveRequest` | `() => boolean \| void \| Promise<boolean \| void>` | Intercept saves; return `true` to continue. |
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

Deprecated: `legacySyncApi` opts into the synchronous hybrid runtime and is removed in 0.6.0; omit it
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
input and delivers bytes through `onSave`, or downloads them when that callback is absent.

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
