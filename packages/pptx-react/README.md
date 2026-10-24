# @betteroffice/pptx-react

React presentation editor with a slide canvas, slide strip, formatting toolbar, and keyboard editing.
The engine runs in a Web Worker by default, with promise-returning reads and exports.

## Install

```bash
bun add @betteroffice/pptx-react react react-dom
```

Peer dependencies: `react` and `react-dom`, version 18 or 19; the PPTX core is a package dependency.

## Usage

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

`fonts` is required. Each entry is a family name and the bytes of a TTF or OTF file, for example
fetched from your server or loaded with `loadBundledFontBytes` from
[`@betteroffice/fonts`](https://www.npmjs.com/package/@betteroffice/fonts).

## Key options

The exported `PptxAsyncEditorProps` type describes the default editor props. All props below are
optional except `fonts`.

| prop | type | description |
| --- | --- | --- |
| `file` | `Uint8Array \| ArrayBuffer` | Presentation bytes to open. |
| `fonts` | `ReadonlyArray<PptxFontFace>` | Font families and bytes for text layout. |
| `readOnly` | `boolean` | Block user edits while retaining navigation and selection. |
| `initialSlide` | `number` | Initial slide, numbered from 1 and clamped to the deck. |
| `fileName` | `string` | Download name; defaults to `presentation.pptx`. |
| `className` | `string` | Class for the editor container. |
| `style` | `CSSProperties` | Inline styles for the editor container. |
| `initialZoom` | `number` | Initial zoom scale; omitted, the slide fits the stage. |
| `loadingIndicator` | `ReactNode` | Shown while the presentation opens. |
| `i18n` | `PptxAsyncEditorProps['i18n']` | UI translations from `@betteroffice/pptx-i18n`. |
| `onReady` | `(api: PptxAsyncEditorApi) => void \| (() => void)` | Receive the async API; optionally return cleanup. |
| `onChange` | `(snapshot: DeckSnapshot) => void` | Receive deck snapshots after changes. |
| `onDocumentChange` | `(change: PptxDocumentChange) => void` | Receive `{ version }` after a committed edit, alongside `onChange`. |
| `onError` | `(error: Error) => void` | Receive loading and editing errors. |
| `onSave` | `(bytes: Uint8Array) => void` | Receive saved bytes; omitted, saving downloads the file. |
| `downloadOnSave` | `boolean` | Whether saving also downloads a copy; defaults to `true` only without `onSave`. |
| `onSaveRequest` | `() => boolean \| void \| Promise<boolean \| void>` | Intercept toolbar and keyboard saves; return `true` to continue. |
| `toolbar` | `ReactNode` | Replace built-in controls; `null` removes them. |
| `showToolbar` | `boolean` | Show the toolbar region; defaults to `true`. |
| `collaboration` | `PptxAsyncCollaborationOptions` | Shared seed, replica callback, and presence options. |
| `plugins` | `readonly PptxPlugin[]` | Host-owned tools, panels, and overlays. |
| `pluginGrants` | `Readonly<Record<string, PptxPluginGrant>>` | Permissions keyed by plugin ID. |
| `onPluginError` | `(error: PptxPluginError) => void` | Receive isolated plugin failures. |

`PptxFontFace` and `DeckSnapshot` are exported by `@betteroffice/pptx`; editor and plugin types are
exported by `@betteroffice/pptx-react`. `ReactNode` is a React type.

Deprecated: `legacySyncApi` opts into the synchronous hybrid runtime and is removed in 0.6.0;
omit it to use the default async editor. Also omit the deprecated `experimentalWorkerOpen` prop.

## Editor API

`onReady` receives `PptxAsyncEditorApi`. Use `saveAsync()` to serialize the presentation and await
`handleAsync()` for asynchronous document access, including `snapshot()`, `exportStructured()`, and
`exportMarkdown()`. Reads and exports through that access object return promises; the API also has
`exportStructured()` and `exportMarkdown()` directly. Work that cannot run against its input or
target rejects with `PptxCommandAdmissionError`.

`flushPendingInput()` waits for accepted input. `goToSlide(slide)` and `selectText(target)` return
promises; slide numbers are 1-based and text offsets use UTF-16. `getPositionAtPoint(clientX, clientY)`
asynchronously returns a hit on the current slide or `null` outside slide content.

Finish pointer gestures before workflows that flush input. Keep the API tied to the mounted
presentation; replacing or closing the document invalidates its access.

## Saving

The toolbar and Ctrl/Cmd+S invoke `onSaveRequest` before serialization. Its result is awaited:
`true` continues saving, while `false` or `void` handles or cancels the request. If saving continues,
`onSave` receives the bytes, and the editor downloads the file using `fileName` when `downloadOnSave`
is `true` (the default without `onSave`).

Calling `saveAsync()` directly returns bytes without invoking those save callbacks.

## Toolbar

The toolbar composition API is experimental. `api.commands` is shared by the built-in controls and
host toolbar parts. Replace controls with
`toolbar`, or hide the region with `showToolbar={false}`.
See [Compose the PPTX toolbar](https://docs.betteroffice.dev/docs/javascript#compose-the-pptx-toolbar)
for composition, external controls, command state, and overflow.

Deprecated: prop-based `EditorToolbar` and `Toolbar` usage is replaced by command-mode controls;
use `EditorToolbar` with `mode="commands"`.

## Plugins

`definePptxPlugin` defines tools installed through `plugins`, with permissions supplied by
`pluginGrants`. This plugin API is experimental and may change in minor releases.
See [Host PPTX editor plugins](https://docs.betteroffice.dev/docs/javascript#host-pptx-editor-plugins)
for panels, overlays, lifecycle, permissions, and error handling.

## Version-checked edit batches

`readContent`, `findText`, `validateEdits`, and `applyEdits` on the editor API return promises and run
after pending input. See [Edit PPTX in version-checked batches](https://docs.betteroffice.dev/docs/javascript#edit-pptx-in-version-checked-batches)
for requests, atomicity, and typed refusals.

## Review agent proposals

Use the access returned by `handleAsync()` to call `propose`, then await `refreshProposals()`.
Pending proposals appear on the canvas with text differences, shape bounds, and speaker-note diffs.
The review controls preview, accept, or reject a group; acceptance creates one undo step.

Escape or **Edit slide** exits review mode. Saving, PNG export, and presentation mode use committed
content. Pending proposals stay in the session; accepted edits are saved and synchronized.

## Collaboration

Pass `collaboration` to co-edit a deck; `onReplica` receives an async replica or `null` on cleanup.
Set `roomId` when switching rooms and seed peers from the same update. Omit `clientId` to generate
one for each opened session.
See the [collaboration guide](https://docs.betteroffice.dev/docs/collaboration) for transport setup.

## Browser integration

Mount the editor in the browser with a container of explicit height. Frameworks that render on the
server should load the editor client-side with server rendering disabled.

## Documentation

- [JavaScript guide](https://docs.betteroffice.dev/docs/javascript)
- [Collaboration guide](https://docs.betteroffice.dev/docs/collaboration)
- [Package guide](https://docs.betteroffice.dev/docs/packages)
- [PPTX core: @betteroffice/pptx](https://www.npmjs.com/package/@betteroffice/pptx)
- [UI translations: @betteroffice/pptx-i18n](https://www.npmjs.com/package/@betteroffice/pptx-i18n)
- [Repository](https://github.com/openooxml/betteroffice)
