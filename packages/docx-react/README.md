# @betteroffice/docx-react

React chrome for the BetterOffice DOCX editor — wraps
[`@betteroffice/docx`](https://www.npmjs.com/package/@betteroffice/docx) in a
drop-in `<DocxEditor>` component with the toolbar, ruler, selection, keyboard,
comments, and tracked-changes UI wired up. Rendering and layout run in the
core's Rust/WebAssembly engine; pages are painted onto canvas.

<!-- TODO(author): add a screenshot/GIF here once hosted; an <img> with an unresolvable src renders broken on npm -->

```bash
bun add @betteroffice/docx-react @betteroffice/docx react react-dom
```

`react` and `react-dom` (18 or 19) are peer dependencies.

## View a document

```tsx
import { DocxEditor } from '@betteroffice/docx-react';
import '@betteroffice/docx-react/styles.css';

<DocxEditor documentBuffer={buffer} mode="viewing" />;
```

`documentBuffer` accepts an `ArrayBuffer`, `Uint8Array`, `Blob`, or `File`.

## Edit a document

```tsx
import { useState } from 'react';
import { DocxEditor } from '@betteroffice/docx-react';
import '@betteroffice/docx-react/styles.css';

export function App() {
  const [file, setFile] = useState<ArrayBuffer>();

  return (
    <>
      <input
        type="file"
        accept=".docx"
        onChange={async (e) => {
          const f = e.target.files?.[0];
          if (f) setFile(await f.arrayBuffer());
        }}
      />
      <DocxEditor
        documentBuffer={file}
        onSave={(bytes) => console.log(`saved ${bytes.byteLength} bytes`)}
      />
    </>
  );
}
```

Without `onSave`, File > Save downloads the edited bytes.

Key props: `documentBuffer` (or a parsed `document`), `onSave`, `onChange`,
`author`, `mode` (`editing` / `suggesting` / `viewing`), `showToolbar`, `toolbar`,
`showRuler`, `showZoomControl`, `showHiddenText`, `i18n`, `measurementFontProvider`. The `ref`
exposes the full editor API (selection, formatting, find/replace, comments,
revisions) and the editor's command store, `commands`.

Vanished (hidden) text stays out of the layout by default; pass
`showHiddenText` to reveal it with normal wrapping.

## What works today

- Editing with Word-faithful pagination; layout runs in Rust, never in the DOM
- Suggesting mode: tracked changes with accept/reject review UI
- Comment threads with replies and resolution, controllable from the host
- Find and replace, headers and footers, footnotes, images, tables
- Zoom control and ruler
- Composable chrome: arrange built-in controls with their state, shortcuts and
  restrictions, add host actions, or drive a toolbar outside the editor
- Host-owned plugins with panels, overlays, sidebar items, lifecycle events,
  and explicitly granted commands and edit batches
- Localized UI via the `i18n` prop
  ([`@betteroffice/docx-i18n`](https://www.npmjs.com/package/@betteroffice/docx-i18n))
- Real-time collaboration with people or agents; the document is a CRDT
- Live collaborator cursors and selections, shown in each peer's color

For font metrics closer to Word, configure `@betteroffice/fonts` once at
module scope, or pass a provider through `measurementFontProvider`. The CDN
entry loads Latin and Japanese coverage fonts without bundling font binaries:

```ts
import { configureDefaultFonts } from '@betteroffice/docx-react';

configureDefaultFonts({ load: () => import('@betteroffice/fonts/cdn') });
```

For offline assets, import `* as fonts` from `@betteroffice/fonts` and call
`configureDefaultFonts({ fonts })`. Without a provider, pagination uses
approximate fallback metrics. CJK coverage fonts are substitutes and can differ
from Word's fonts; see [font configuration](../fonts/README.md).

## Collaboration

The document is a CRDT — pass a `collaboration` prop and wire a transport to
co-edit with other people or an agent. `onReplica` hands you the session as a
`CollaborationReplica`; drive it with a `CollaborationProvider` over any
transport (a WebSocket relay, etc.). Every window must boot from the same shared
state, so pass identical `initialUpdate` bytes to each peer.

```tsx
import { CollaborationProvider } from '@betteroffice/docx/collaboration';

<DocxEditor
  documentBuffer={file}
  collaboration={{
    clientId,
    initialUpdate: sharedSeed, // identical bytes for every peer
    onReplica: (replica) => {
      if (!replica) return;
      const provider = new CollaborationProvider(replica, transport, {
        user: { name: "Ada" }, // shown on this peer's remote caret
      });
      provider.connect();
    },
  }}
/>;
```

## Host saves and pending input

`onSaveRequest` owns File > Save and Cmd/Ctrl+S before serialization. Its awaited
callback can perform locks, revision checks, and persistence. Returning `true`
continues the built-in export/download; returning `false` or nothing suppresses it.
Overlapping UI save requests share one workflow. Errors reach `onError`.

`DocxEditorRef.flushPendingInput()` and `PagedEditorRef.flushPendingInput()` wait
until input accepted before the call and its selection are authoritative in Yrs.
They wait for active IME composition to end, and reject on input failure, unmount,
or document replacement. A failed input queue remains failed until a new session
is loaded. The promise does not wait for browser painting.

```tsx
<DocxEditor
  ref={editorRef}
  onSaveRequest={async () => {
    await validateRevision();
    await editorRef.current!.flushPendingInput();
    const bytes = await editorRef.current!.save();
    if (bytes) await persistDocument(bytes);
  }}
/>
```

`save()` flushes input and exports directly, so calling it inside `onSaveRequest`
does not invoke that callback again. `onSave(buffer)` remains the notification
after export. Before direct session reads or mutations, await `flushPendingInput()`;
after a mutation, use `syncYrsInputState(true)` to refresh the editor.

## Compose the toolbar

Every built-in control runs through one command store, `ref.commands`, which
hosts can use for their own chrome. This command and toolbar composition API is
experimental and may change in minor releases. Pass `toolbar` to replace the
default chrome with an arrangement of public parts; the children of
`EditorToolbar.Toolbar` are the complete row, in your order:

```tsx
import {
  DocxEditor,
  EditorToolbar,
  ToolbarButton,
  ToolbarCommandButton,
  ToolbarCommandSelect,
  ToolbarGroup,
} from '@betteroffice/docx-react';

function CompactToolbar({ onShare }: { onShare(): void }) {
  return (
    <EditorToolbar>
      <EditorToolbar.Toolbar>
        <ToolbarGroup label="Formatting">
          <ToolbarCommandSelect id="paragraphStyle" />
          <ToolbarCommandButton id="bold" />
        </ToolbarGroup>
        <ToolbarCommandButton id="undo" />
        <EditorToolbar.Review />
        <ToolbarButton title="Share" onClick={onShare}>Share</ToolbarButton>
      </EditorToolbar.Toolbar>
    </EditorToolbar>
  );
}

<DocxEditor documentBuffer={file} toolbar={<CompactToolbar onShare={share} />} />;
```

`toolbar` omitted renders the default chrome, `null` renders none, and
`showToolbar={false}` hides either. The default chrome is hidden for `readOnly`;
chrome you supply still renders, with its writing controls disabled.

To place the toolbar outside the editor, capture the ref in state and provide it.
Until the editor attaches, `commands={null}` reports every command unavailable:

```tsx
const [commands, setCommands] = useState<DocxCommandStore | null>(null);

<DocxCommandProvider commands={commands}>
  <CompactToolbar onShare={share} />
</DocxCommandProvider>
<DocxEditor ref={(editor) => setCommands(editor?.commands ?? null)} toolbar={null} />
```

Outside the editor, `EditorToolbar` supplies its own styling root and the
editor's locale, and its keyboard shortcuts reach that editor only.

- **Parts.** `ToolbarCommandButton` binds a command (parameterized ones need
  `args`, such as `{ value: 'center' }` for `alignment`), `ToolbarCommandSelect`
  renders a selector's built-in picker, `ToolbarCommand` renders any command's
  default control, and `ToolbarButton`, `ToolbarGroup` and `ToolbarSeparator`
  hold host actions. `useDocxCommand(id)` and `useDocxCommandState(id)` bind
  custom controls.
- **State.** `getState(id, args?)` returns `{ enabled, active, value, options }`;
  a disabled command always carries `disabledReason: { code, message }`, and
  controls expose that message as their accessible description. Marks report
  `'mixed'` for mixed selections.
- **Results.** `execute(id, args)` resolves to `executed`, `noop`, `opened` (a
  dialog or picker), `requested` (handed to the host, such as a controlled mode
  or `onSaveRequest`), or `{ ok: false, failure }`. Commands run after input
  accepted before the call and check availability again first, so a stale
  enabled state never authorizes a change. A dialog or picker a command opens
  applies to the document and selection it opened with, or fails with
  `document-replaced` or `target-changed`; print waits until accepted input is
  rendered.
- **Modes.** `readOnly` and viewing mode refuse writes (`read-only`,
  `viewing-mode`); navigation, find, zoom, print and save remain available.
  Suggesting mode keeps direct character formatting untracked, tracks paragraph
  styles and row edits, and refuses operations it cannot record as suggestions
  (`suggesting-unsupported`), such as inserting tables or breaks.
- **Focus and overflow.** Pointer clicks keep the document focused; keyboard
  activation keeps focus in the toolbar, and dialogs take focus. At narrow
  widths trailing groups move into a More menu with arrow, Home/End, typeahead
  and submenu navigation, keeping every choice of the built-in controls; custom
  colors and sizes are asked for in a dialog. Wrap host content in
  `ToolbarOverflow` to give it a menu entry; content without one, and any group
  holding it, stays in the row.

## Host edit batches

`DocxEditorRef.readParagraphs()`, `findText()`, `validateEdits()` and `applyEdits()`
flush pending input first, then run the version-checked batch API of
[`@betteroffice/docx`](https://www.npmjs.com/package/@betteroffice/docx) on the
open document. Build `expectVersion` and targets from a flushed read:

```tsx
const read = await editorRef.current!.readParagraphs({ view: 'accepted' });
if (!read.ok) return;
const result = await editorRef.current!.applyEdits({
  expectVersion: read.version,
  steps: [
    {
      op: 'insertParagraphs',
      target: { story: 'body', paraId: read.paragraphs[0].paraId },
      at: 'end',
      paragraphs: [{ text: 'Inserted by the host', styleId: 'Heading2' }],
    },
  ],
});
```

A batch commits as one undo step and refreshes every story it changed, or returns
a typed refusal. A refusal never rolls back typing that the flush committed, so
typing that lands between the read and the batch refuses with `stale-version`.
Read-only editors refuse with `read-only` (`locked-target` stays for locked
content); in suggesting mode every step must carry `suggest`. The promise rejects when the document is replaced while input
flushes. `proposeChange`, `addComment` and `applyFormatting` resolve their
`{ paraId, search }` targets through the same Rust resolver in the accepted view.

## Host plugins

The plugin API is experimental and may change in minor releases.

Host-owned tools (review aids, templates, checks) install through the `plugins`
prop. A plugin contributes a panel, an overlay, sidebar cards, and commands, and
works through restricted clients rather than the editor or its session:

```tsx
import { DocxEditor, defineDocxPlugin } from '@betteroffice/docx-react';

type State = { version: string | null; count: number };

const review = defineDocxPlugin<State>({
  id: 'acme.review',
  createState: () => ({ version: null, count: 0 }),
  async onEvent(context, event) {
    if (event.type !== 'load' && event.type !== 'document-change') return;
    const read = await context.read.readParagraphs({ view: 'accepted' });
    if (read.ok) {
      context.setState({ version: read.version, count: read.paragraphs.length }, read.version);
    }
  },
  panel: {
    title: 'Review',
    placement: 'right',
    render: ({ context }) => <p>{context.state.count} paragraphs</p>,
  },
});

<DocxEditor
  documentBuffer={file}
  plugins={[review]}
  pluginGrants={{ 'acme.review': { document: 'write', editBatches: true } }}
  onPluginError={(error) => report(error)}
/>;
```

- **Lifecycle.** Once a document is ready, each plugin gets fresh state,
  `initialize`, then one `load` event (`loaded`, `replaced`, or `attached` for a
  plugin added to an open document), and its contributions appear. A document
  change the hook did not make itself aborts it and delivers `load` again; after
  ten such runs the plugin is stopped and reported. It then receives
  `document-change` (the committed version only, for typing, remote edits, undo,
  commands and batches, its own included, never for refusals or no-ops),
  `selection-change`, `mode-change` (with the effective `readOnly`),
  `layout-change` and `grants-change`. Events describe current state: several
  changes may arrive as one, and a newer one aborts the hook still handling the
  previous (`context.signal`), except a change that hook's own edit batch
  made. Replacing the document, removing the plugin,
  changing its `revision`, unmounting, or a failure ends the activation: its
  signals abort, its clients refuse, and every `onCleanup` disposer runs once
  with the reason. A disposer registered afterwards runs immediately. Plugins
  are matched by `id` and `revision`, so new array or callback identities and
  reordering keep their state. `context.run(action)` gives event handlers a
  fresh context and isolates their failures.
- **Stale results.** `setState` returns false once the context is superseded or
  ended, or when the document is no longer at `atVersion` (by default the
  version the context was created at). Pass the version a read returned so an
  older read never overwrites newer state.
- **Grants.** Without a grant a plugin reads, validates and navigates only, and
  `context.edits` is null. Built-in commands need their id in `commands`, and
  mutating ones also `document: 'write'`; edit batches need `document: 'write'`
  and `editBatches`, and `history: 'none'` also `untrackedHistory`. Grants,
  editor mode and document policy are checked again right before each change,
  so a revoked grant, viewing mode, `readOnly` or a replaced document refuses
  even through a client obtained earlier. Mutating built-in commands have no
  authoritative lock policy yet and refuse plugins with `unsupported-policy`;
  plugins change documents through edit batches, whose Rust policy refuses
  locked content. Suggesting mode follows each operation's own rules; a batch
  step needs `suggest` there.
- **Contributed commands** register as `plugin:<pluginId>/<id>` on
  `ref.commands`. They always run with their own plugin's clients, even when the
  host's toolbar, shortcuts or `ref.commands` invoke them, and outside the input
  queue, so a handler can await its own batches. `execute` returns
  `{ ok: true, status }` or a failure with the plugin's own code, or a refused
  edit batch as-is; callers receive it unchanged. `mutatesDocument` disables a
  command in viewing and read-only modes but grants nothing. Plugin shortcuts
  must use Mod or Alt, or a function key; built-in and clipboard shortcuts win,
  and a clashing plugin shortcut is not bound and is reported. Keys pressed in
  plugin chrome are the user's: built-in shortcuts act on the editor there, and
  a plugin's text fields keep their own editing keys. `toolbar` lists
  local ids in order: the default toolbar shows them, and replacement chrome
  places `<DocxPluginToolbar />` where it wants them. `ToolbarCommandButton`,
  `ToolbarCommand`, `useDocxCommand` and `useDocxCommandState` also take a
  contributed id; a button renders nothing while no plugin contributes it.
- **Panels** dock left, right or bottom of the document, beside the toolbar
  rather than inside the scroll area, with tabs when several share a side.
  `preferredSize` is clamped to 40% of the editor. In a narrow editor side docks
  show their tabs, and a tab opens its panel as a drawer over the document that
  Escape closes. Collapsing a panel keeps the plugin running.
- **Geometry.** `context.geometry` is null until a rendered layout shows the
  current version, and whenever it falls behind. `geometry.dom` answers in
  pages-container units divided by zoom; `geometry.toOverlayRect(rect)` converts
  one of those rectangles into pixels of the unscaled overlay layer, once, or
  returns null after that layout stops being rendered. `geometry.dom` is
  experimental and may be replaced by a data-only facade. The layer ignores the
  pointer; interactive overlay elements set `pointer-events: auto`.
  `snapshot.selection.displayRange` belongs to one layout and is never an edit
  target.
- **Sidebar cards** anchor to `{ version, story, paraId }` and show only while
  the document is at that version and the body paragraph resolves uniquely.
  `render` receives the `item` it draws, so one component can draw every card;
  a card keeps its React state while the plugin returns its `id`, and starts
  afresh for a new id or document. Their ids are namespaced, so they never
  collide with comments.
- **Navigation.** `navigation.scrollToParagraph(target, { expectVersion })`
  flushes input, waits briefly for a layout of that version, and keeps focus
  and selection unless `focus: true`. It succeeds only when it scrolled, and
  otherwise says why (`stale-version`, `missing-target`, `ambiguous-target`,
  `layout-unavailable`, `unsupported`).
- **Failures.** Every contribution renders behind its own error boundary with a
  command context restricted to its plugin, so public command hooks inside it
  act for the plugin. A throwing hook, renderer, sidebar or command-state
  function, command, or `context.run` action stops only that plugin, runs its
  cleanups and reports `{ pluginId, generation, phase, error }` to
  `onPluginError`; it stays stopped until a new document, `revision` or
  reinstallation. Plugins run in the page's realm: grants govern the supported
  API, not a sandbox.
- **Legacy props.** `pluginOverlays` and `pluginSidebarItems` still render as
  unmanaged content, below managed overlays and merged with managed cards;
  `pluginRenderedDomContext` is ignored while plugins are installed, and
  `onRenderedDomContextReady` receives the editor's own context. The
  snapshot-based `EditorPluginCore` types in `@betteroffice/docx/plugin-api`
  are deprecated in favour of this API.

## Framework notes

Import `@betteroffice/docx-react/styles.css` once (in a bundler entry or, under
Next.js, at the page/layout level — CSS imported inside a `next/dynamic`
component does not attach in production builds). The editor is browser-only
(canvas, wasm, workers); under Next.js load it with `next/dynamic` and
`ssr: false`.

[JavaScript guide](https://docs.betteroffice.dev/docs/javascript) ·
[Changelog](https://github.com/openooxml/betteroffice/blob/main/packages/docx-react/CHANGELOG.md) · Apache-2.0.
