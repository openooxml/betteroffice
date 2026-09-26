# @betteroffice/pptx-react

React chrome for the BetterOffice PPTX editor — wraps
[`@betteroffice/pptx`](https://www.npmjs.com/package/@betteroffice/pptx) in a
slide canvas, slide strip, formatting toolbar, and keyboard editing surface.
Parsing, layout, and text shaping run in the core's Rust/WebAssembly engine;
slides are painted onto canvas.

<!-- TODO(author): add a screenshot/GIF here once hosted; an <img> with an unresolvable src renders broken on npm -->

```bash
bun add @betteroffice/pptx-react @betteroffice/pptx react react-dom
```

`react` and `react-dom` (18 or 19) are peer dependencies.

## Render a presentation

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

Font bytes are supplied by the host and registered with the Rust shaper; pair
with [`@betteroffice/fonts`](https://www.npmjs.com/package/@betteroffice/fonts)
for metric-compatible open faces.

## Edit

Click text to place the Rust-computed caret, type to edit the yrs story and
trigger Rust reflow, drag or resize shapes on the canvas, or use the toolbar
for bold, italic, size, color, slides, text boxes, image insertion, and shape ordering.

Props: `file`, `fonts`, `collaboration`, `i18n`, `className`, `fileName`,
`onReady` (exposes the core `PresentationHandle`, the editor's `commands`, a
`refresh` callback for host-driven edits, `refreshProposals`, and `save`),
`onChange` (deck snapshots), `onError`, `onSave` (receives the saved bytes;
without it, saving downloads the file), `readOnly`, `toolbar`, and `showToolbar`.

## Compose the toolbar

Every built-in control runs through one command store per editor,
`api.commands`. Hosts arrange the same controls with their own actions:

```tsx
import {
  EditorToolbar,
  PptxEditor,
  ToolbarButton,
  ToolbarCommandButton,
  ToolbarCommandSelect,
  ToolbarGroup,
} from '@betteroffice/pptx-react';

<PptxEditor
  file={file}
  fonts={fonts}
  toolbar={
    <EditorToolbar mode="commands">
      <EditorToolbar.Toolbar>
        <ToolbarGroup label="Text">
          <ToolbarCommandSelect id="fontFamily" />
          <ToolbarCommandButton id="bold" />
        </ToolbarGroup>
        <ToolbarCommandButton id="undo" />
        <ToolbarCommandButton id="slideshow" />
        <ToolbarButton title="Share" onClick={share}>Share</ToolbarButton>
      </EditorToolbar.Toolbar>
    </EditorToolbar>
  }
/>;
```

- **Toolbar region.** `showToolbar={false}` hides it; otherwise omitting
  `toolbar` keeps the default controls (read-only shows only the Present
  button), `null` removes them, and supplied chrome replaces them, also while
  `readOnly`. Collaborator chips stay beside supplied chrome.
- **Outside the editor.** Capture `api.commands` from `onReady` into state and
  render the parts under `<PptxCommandProvider commands={commands}>`;
  `commands={null}` reports the editor as unavailable until it is ready. Such
  chrome follows the editor's locale, and shortcuts pressed inside it reach the
  editor.
- **Command contract.** `getState(id, args?)` returns serializable state
  (`enabled`, `active` with `'mixed'`, `value`, and `options` that each carry
  their own `state`); a disabled state
  always carries `disabledReason: { code, message }`, for example
  `text-selection-required`, `shape-required`, `z-order-boundary`,
  `review-active` or `proposal-stale`. `getDescriptor(id)` lists the label key
  and shortcuts, which labels show per platform (`Ctrl+B`, `⌘B`).
  `usePptxCommand` and `usePptxCommandState` bind custom controls.
- **Ordering.** `execute(id, args)` waits for input accepted before it, such as
  a picture still decoding, keeps later typing behind it, checks availability
  again and resolves to `executed`, `noop`, `opened`, `requested` or a coded
  failure (`input-failed`, `document-replaced`, `target-changed`,
  `gesture-active`, `command-failed`). Keystrokes typed meanwhile land in the
  text they were typed into, and a picture on the slide the picker opened on;
  input queued for a replaced document is dropped. Save runs the host's `onSaveRequest` outside that queue, so the request
  may await `flushPendingInput()`.
- **Authority.** Commands enforce read-only mode and the other gates for the UI.
  `api.handle` stays direct, unrestricted host access; call `api.refresh()`
  after using it.
- **Overflow.** At narrow widths trailing groups move into a More menu with
  arrow, Home/End, typeahead and submenu navigation. Built-in controls keep all
  their choices there; host buttons and dropdowns get entries automatically,
  and `ToolbarOverflow` gives other content one. A group holding content
  without an entry stays in the row.
- **Compatibility.** The prop-based `EditorToolbar`, `Toolbar`,
  `EditorToolbarContext` and `useEditorToolbar` still work and are deprecated:
  without `mode` they bind to their props, host children follow the default
  controls, and command-mode chrome rejects those props. Inside
  `EditorToolbar mode="commands"`, `useEditorToolbar()` returns a view of the
  commands whose callbacks run them; mixed marks read as `undefined`.

## Host editing controls

`onSaveRequest` runs for toolbar and Ctrl/Cmd+S saves before serialization.
Return `true` to continue built-in saving; `false` or `void` handles or cancels
it. Promises are awaited and concurrent requests are coalesced. `onSave` still
receives the resulting bytes when built-in saving continues. A request waiting
on a replaced or closed document is discarded.

The API received by `onReady` exposes `flushPendingInput(): Promise<void>` and
the `commands` store described above.
Await it before inspecting or mutating the core from a host workflow, then call
`api.save()` for explicit serialization without re-entering `onSaveRequest`.
`save()` remains synchronous and rejects while asynchronous input is pending.
Flush rejects stale document handles, failed input, and unfinished pointer
gestures. Finish or cancel the gesture before retrying.

PPTX keyboard and notes edits are synchronous; flushing also waits for accepted
image imports. A failed import rejects a flush waiting for it and is reported
through `onError`; it does not block later saves of the existing presentation.

`api.getPositionAtPoint(clientX, clientY)` returns a shape or text hit with
`slide` (1-based), `slideId`, and `shapeId`; text hits include `storyId` and the
UTF-16 `position`. It uses the current slide and canvas bounds without changing
focus or selection. Outside content, unavailable geometry, stale handles, and
proposal previews return `null`. Use `api.refresh()` after core mutations.

`api.readContent()`, `api.findText()`, `api.validateEdits()` and
`api.applyEdits()` run the core's version-checked edit batches after flushing
pending input, and `api.version()` returns the flushed session version:

```ts
const read = await api.readContent();
if (read.ok) {
  const result = await api.applyEdits({
    expectVersion: read.version,
    steps: [{ op: 'setSlideNotes', target: { slideId: read.slides[0].id }, text: 'Opening' }],
  });
  if (!result.ok) console.warn(result.failure.code);
}
```

An applied batch refreshes the slides, history and proposals once and calls
`onChange`; refusals and no-op batches refresh nothing. A refusal never rolls
back input the flush committed, so an edit that lands between the read and the
batch refuses with `stale-version`. Read-only editors refuse `validateEdits` and
`applyEdits` with `read-only`, and the promises reject when the presentation is
replaced while input flushes.

## Host plugins

The plugin API is experimental and may change in minor releases.

Host-owned tools (review aids, checks, templates) install through the `plugins`
prop. A plugin contributes a docked panel, an overlay and commands, and works
through restricted clients rather than the editor API or the presentation
handle:

```tsx
import { PptxEditor, definePptxPlugin } from '@betteroffice/pptx-react';

type State = { version: string | null; slides: number };

const review = definePptxPlugin<State>({
  id: 'acme.review',
  createState: () => ({ version: null, slides: 0 }),
  async onEvent(context, event) {
    if (event.type !== 'load' && event.type !== 'document-change') return;
    const read = await context.read.readContent();
    if (read.ok) context.setState({ version: read.version, slides: read.slides.length }, read.version);
  },
  panel: {
    title: 'Review',
    placement: 'right',
    render: ({ context }) => <p>{context.state.slides} slides</p>,
  },
  overlay: ({ context, geometry }) => {
    const target = context.snapshot.selection?.target;
    const box = target && target.kind !== 'slide' ? geometry.getShapeRect(target.shapeId) : null;
    return box ? (
      <div style={{ position: 'absolute', left: box.x, top: box.y, width: box.width, height: box.height, outline: '2px dashed #2563eb' }} />
    ) : null;
  },
});

<PptxEditor
  file={bytes}
  fonts={fonts}
  plugins={[review]}
  pluginGrants={{ 'acme.review': { document: 'write', editBatches: true } }}
  onPluginError={(error) => report(error)}
/>;
```

- **Lifecycle.** Once a presentation is open, each plugin gets fresh state,
  `initialize`, then one `load` event (`loaded`, `replaced`, or `attached` for a
  plugin added to an open presentation), and its contributions appear. A
  presentation change the hook did not make itself aborts it and delivers
  `load` again; after ten such runs the plugin is stopped and reported. It then
  receives `document-change` (the committed version only, for typing, commands,
  batches, its own included, undo, redo and remote updates, never for refusals
  or no-ops), `selection-change`, `mode-change` (with `readOnly`),
  `layout-change` and `grants-change`. Events describe current state: several
  changes may arrive as one, and a newer one aborts the hook still handling the
  previous (`context.signal`), except a change that hook's own edit batch made.
  Replacing the presentation, removing the plugin, changing its `revision`,
  unmounting, or a failure ends the activation: its signals abort, its clients
  refuse, and every `onCleanup` disposer runs once with the reason. Plugins are
  matched by `id` and `revision`, so new array or callback identities and
  reordering keep their state. `definePptxPlugin` copies the panel, commands
  and toolbar, so changing them takes a new definition. `context.run(action)`
  gives event handlers a fresh context and isolates their failures. `onReady` is
  unaffected by plugins.
- **Stale results.** `setState` returns false once the context is superseded or
  ended, or when the presentation is no longer at `atVersion` (by default the
  version the context was created at).
- **Reads and edits.** `context.read` offers `version`, `readContent`,
  `findText` and `validateEdits`, and `context.edits.applyEdits` the editor's
  version-checked batches. Each call waits for pending input in order with it,
  but plugin handlers run outside that queue, so a command can await its own
  batch.
- **Grants.** Without a grant a plugin reads, validates and navigates only, and
  `context.edits` is null. Built-in commands need their id in `commands`, and
  mutating ones also `document: 'write'`; edit batches need `document: 'write'`
  and `editBatches`, and `history: 'none'` also `untrackedHistory`. The grant
  and `readOnly` are checked again right before each change, so a revoked grant,
  `readOnly` or a replaced presentation refuses even through a client obtained
  earlier. Mutating built-in commands have no authoritative lock policy yet and
  refuse plugins with `unsupported-policy`; plugins change presentations through
  edit batches.
- **Contributed commands** register as `plugin:<pluginId>/<id>` on
  `api.commands`. They always run with their own plugin's clients, even when the
  toolbar, a shortcut or the host invokes them. `execute` returns
  `{ ok: true, status }` or a failure with the plugin's own code, or a refused
  edit batch as-is; callers receive it unchanged. `mutatesDocument` disables a
  command while read-only but grants nothing. Plugin shortcuts must use Mod or
  Alt, or a function key; built-in and clipboard shortcuts win, and a clashing
  plugin shortcut is not bound and is reported. `toolbar` lists local ids in
  order: the default toolbar shows them, and replacement chrome places
  `<PptxPluginToolbar />` where it wants them. `ToolbarCommandButton`,
  `ToolbarCommand`, `usePptxCommand` and `usePptxCommandState` also take a
  contributed id.
- **Selection and navigation.** `snapshot.selection` names the current slide
  (`slideId` and the one-based `slide`) and a slide, shape or text target; text
  keeps its `anchor` and `focus` direction in UTF-16 story offsets.
  `navigation.goToSlide`, `selectShape` and `selectText` resolve session ids
  after pending input against `expectVersion`, keep keyboard focus unless
  `focus: true`, and refuse with `stale-version`, `missing-target`,
  `layout-unavailable` or `unsupported` (during proposal review) rather than
  retarget. Ids are session-scoped: they do not survive saving and reopening.
- **Geometry.** `context.geometry` exists only while the canvas shows a painted
  frame of the current version, and not during proposal review.
  `layout.width` and `height` are the unzoomed slide in display-list pixels and
  `layout.zoom` the resolved scale, fit included. `geometry.toOverlayRect` takes
  a `slide-emu` (9,525 EMU per pixel) or `slide-px` rectangle and returns pixels
  of the unscaled overlay layer, which sits on the slide canvas below selection
  handles and proposal controls and ignores the pointer unless an element sets
  `pointer-events: auto`. `getShapeRect(shapeId)` returns the rendered bounds of
  a shape and its group descendants, and `getPositionAtPoint` adds the layout's
  `version` and `id` to a hit. Every method returns null once its layout is gone.
- **Panels** dock left, right or bottom of the slide, beside the thumbnail rail
  and outside the slide's keyboard handling, with tabs when several share a
  side. `preferredSize` is clamped to 40% of the workspace. In a narrow editor
  side docks show their tabs, and a tab opens its panel as a drawer that Escape
  closes. Collapsing a panel keeps the plugin running.
- **Failures.** Every contribution renders behind its own error boundary with a
  command context restricted to its plugin. Keys pressed in plugin chrome are
  the user's: built-in shortcuts act on the editor there, and a plugin's text
  fields keep their own editing keys. A throwing hook, renderer,
  command-state function, command or `context.run` action stops only that
  plugin, runs its cleanups and reports `{ pluginId, generation, phase, error }`
  to `onPluginError`. Plugins run in the page's realm: grants govern the
  supported API, not a sandbox.

## What works today

- Slide rendering with Rust layout and text shaping, painted onto canvas
- Canvas interactions: shape selection, drag, and resize
- Text editing with caret and selection computed by the engine
- Slide management (add, delete), text boxes, image insertion, shape ordering, undo/redo
- Saving the deck back to `.pptx` — toolbar button, Ctrl/Cmd+S, or `api.save()`
- A command store behind every control, with composable toolbar parts and an
  accessible overflow menu
- Host-owned plugins with panels, overlays, lifecycle events, and explicitly
  granted commands and edit batches
- Localized UI via the `i18n` prop
  ([`@betteroffice/pptx-i18n`](https://www.npmjs.com/package/@betteroffice/pptx-i18n))
- Real-time collaboration with people or agents; the deck is a CRDT
- Live collaborator shape selections and presence chips, shown in each peer's color
- Agent proposal review directly on the slide canvas, with inline text diffs,
  old/new shape bounds, before/after previews, acceptance, rejection, and Undo

## Review agent proposals

Use the editor API supplied to `onReady` to stage edits from your agent, then
refresh the pending list:

```ts
api.handle.propose('editor-agent', 'Clarify the speaker notes', [{
  type: 'setSlideNotes',
  slideId: api.handle.snapshot().slides[0].id,
  text: 'Explain the customer outcome before the implementation details.',
}]);
api.refreshProposals();
```

Pending edits appear directly on the slide canvas: deleted text is red and
struck through, inserted text is green and underlined, and moved/resized shapes
show their previous and proposed bounds. Speaker notes have their own inline
text diff. The canvas toolbar selects a proposal when several affect the slide
and provides acceptance, rejection, and access to review details. Acceptance is
available after the current diff has painted successfully.

The canvas is in review mode while showing a diff. **Edit slide** (or Escape)
returns to normal editing; **Show changes** restores the diff. This keeps
temporary review offsets separate from editable text. Saving, PNG export, and
presentation mode use the actual document. Changes from other users refresh the
review, and stale targets require the explicit review described below.

The **Agent proposals** button shows the pending count. Each group shows its
author, rationale, targets, and text changes. **Preview** opens current and
proposed slide renderings, with a target selector for groups spanning multiple
shapes or slides. Notes changes also appear as text because notes are outside
the slide canvas.

Accepting a group updates the editor, calls `onChange`, and creates one Undo
step. Rejecting it leaves the deck untouched. If a target changed, preview its
current state before choosing **Apply updated proposal**. A further target
change detected at that click refreshes the preview for another review.

Pending proposals are session-local and disappear when the deck closes. Only
accepted edits are saved and synchronized. Existing host-driven edits may keep
using `api.refresh()`, which also refreshes the proposal list.

## Collaboration

Pass a `collaboration` prop to co-edit a deck live. `onReplica` hands you the
session; drive it with a `CollaborationProvider` over any transport. Every peer
must boot from the same `initialUpdate` seed.

```tsx
import { CollaborationProvider } from '@betteroffice/pptx';

<PptxEditor
  file={file}
  fonts={fonts}
  collaboration={{
    clientId,
    initialUpdate: sharedSeed,
    onReplica: (replica) => {
      if (!replica) return;
      const provider = new CollaborationProvider(replica, transport, {
        user: { name: "Ada" }, // shown on this peer's presence chip
      });
      provider.connect();
    },
  }}
/>;
```

## Framework notes

The editor is browser-only (canvas, wasm); under Next.js load it with
`next/dynamic` and `ssr: false`.

[JavaScript guide](https://docs.betteroffice.dev/docs/javascript) ·
[Changelog](https://github.com/openooxml/betteroffice/blob/main/packages/pptx-react/CHANGELOG.md) · Apache-2.0.
