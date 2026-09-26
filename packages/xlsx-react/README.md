# @betteroffice/xlsx-react

React chrome for the BetterOffice XLSX editor — wraps
[`@betteroffice/xlsx`](https://www.npmjs.com/package/@betteroffice/xlsx) in a
drop-in `<XlsxEditor>` component with the toolbar, selection, keyboard, and
clipboard wired up. Formula calculation and rendering run in the core's
Rust/WebAssembly engine; the grid is painted onto canvas.

<!-- TODO(author): add a screenshot/GIF here once hosted; an <img> with an unresolvable src renders broken on npm -->

```bash
bun add @betteroffice/xlsx-react @betteroffice/xlsx react react-dom
```

`react` and `react-dom` (18 or 19) are peer dependencies.

## Render a workbook

```tsx
import { XlsxEditor } from "@betteroffice/xlsx-react";

<XlsxEditor file={bytes} fileName="report.xlsx" />;
```

`file` is a `Uint8Array` of `.xlsx` bytes; omit it to render an empty frame.

## Open and save

```tsx
import { useState } from "react";
import { XlsxEditor } from "@betteroffice/xlsx-react";

export function App() {
  const [file, setFile] = useState<Uint8Array>();

  return (
    <>
      <input
        type="file"
        accept=".xlsx"
        onChange={async (e) => {
          const f = e.target.files?.[0];
          if (f) setFile(new Uint8Array(await f.arrayBuffer()));
        }}
      />
      <XlsxEditor
        file={file}
        fileName="workbook.xlsx"
        onSave={(bytes) => console.log(`saved ${bytes.length} bytes`)}
      />
    </>
  );
}
```

Without `onSave`, the save button downloads the edited bytes.

Props: `file`, `fileName`, `onSave`, `onChange`, `onReady` (a handle for
host/agent-driven edits and the editor's `commands`), `collaboration`, `i18n`,
`readOnly`, `toolbar`, `showToolbar`, and `className`.

## What works today

- Cell editing with formula recalculation of dependents on every edit
- Editing toolbar: number formats, fonts, colors, borders, alignment, merges
- Composable toolbar: built-in controls, the formula bar and host actions in your
  own order, bound to one command store
- Agent proposals: in-cell tracked-change ghosts plus an accept/reject panel
- Version-checked edit batches through `onReady`, after pending input commits
- Host-owned plugins with panels, overlays, lifecycle events, and explicitly
  granted commands and edit batches
- TSV clipboard copy/paste
- Accessible grid mirroring the painted canvas for screen readers
- Localized UI via the `i18n` prop
  ([`@betteroffice/xlsx-i18n`](https://www.npmjs.com/package/@betteroffice/xlsx-i18n))
- Real-time collaboration with people or agents; the workbook is a CRDT
- Live collaborator selections and presence chips, shown in each peer's color

## Compose the toolbar

Every built-in control runs through the editor's command store. This command
and toolbar composition API is experimental and may change in minor releases.
Replace the default toolbar with the parts you need, in your order, next to your
own actions:

```tsx
import {
  EditorToolbar,
  ToolbarButton,
  ToolbarCommandButton,
  ToolbarCommandSelect,
  ToolbarGroup,
  XlsxEditor,
} from "@betteroffice/xlsx-react";

<XlsxEditor
  file={file}
  toolbar={
    <EditorToolbar mode="commands">
      <EditorToolbar.Toolbar>
        <ToolbarGroup label="Formatting">
          <ToolbarCommandSelect id="numberFormat" />
          <ToolbarCommandButton id="bold" />
        </ToolbarGroup>
        <ToolbarCommandButton id="undo" />
        <ToolbarButton title="Share" onClick={share}>Share</ToolbarButton>
      </EditorToolbar.Toolbar>
      <EditorToolbar.FormulaBar />
    </EditorToolbar>
  }
/>;
```

- `toolbar` omitted renders the default toolbar (hidden when `readOnly`), `null`
  renders none, and supplied chrome renders in its place, also when read-only.
  `showToolbar={false}` hides the region either way.
- `<EditorToolbar.Toolbar>` children are the complete arrangement in command mode;
  without children it renders the default controls. `EditorToolbar.FormulaBar`
  is the editor's name box and formula input and renders only inside the editor.
- `ToolbarCommand` renders a command's built-in control, `ToolbarCommandButton` a
  button (bind `args`, such as `{ value: "currency" }` for `numberFormat`), and
  `ToolbarCommandSelect` a selector's picker. `useXlsxCommand(id, args)` and
  `useXlsxCommandState(id, args)` bind custom controls.
- Outside the editor, capture `api.commands` from `onReady` into state and wrap
  the same parts in `<XlsxCommandProvider commands={commands}>`; pass `null`
  until the editor is ready. Its shortcuts then reach that editor.
- `commands.getState(id, args?)` returns serializable state (`enabled`,
  `active`, `value`, and `options` whose entries carry their own `state`); a
  disabled command always carries `disabledReason` with a `code` and a
  localized `message`.
- `commands.execute(id, args)` runs in order with pastes, cell entries and
  chart moves accepted before it: it ends an IME composition, writes the text
  typed so far, then checks availability again. It resolves to
  `{ ok: true, status }` (`executed`, `noop`, `opened`, `requested`) or
  `{ ok: false, failure }`, for example `input-failed`, `document-replaced`,
  `target-changed` when the selection moved while it waited, `gesture-active`
  during a chart drag, `proposal-stale`, or `command-failed` when the change
  could not be applied. A cell entry the workbook refuses
  stays open at its own cell for correction; until it is corrected or discarded
  with Escape, every command fails with `input-failed`. Print waits until the
  canvas has painted the written text, and fails with `render-failed` when it
  cannot.
- `api.save()` stays synchronous and never returns bytes without accepted
  input: it throws `XlsxSaveRefusedError` with `code` `input-pending` while
  entries or a paste still wait to be written, and `input-failed` while a
  refused entry waits for correction. `await api.commands.execute("save", null)`
  waits for that input and then saves through `onSave`. The API's version, read
  and batch methods reject with `XlsxCommandAdmissionError` instead (see
  [Edit batches](#edit-batches)); both errors are exported to branch on `code`.
- Every cell write goes through one queue. `api.selectCells`, `api.clearSelection`
  and switching sheets close the open entry and write it at once when nothing
  waits; otherwise it is queued behind the earlier input, so it may not have
  landed when they return. Commands run right after them, even in the same
  handler, act on the new selection and sheet.
- The store enforces read-only mode, the selection a command needs and its
  arguments for the editor's UI. `api.handle` stays unrestricted host
  authority, and the store is not workbook protection.
- Narrow toolbars move trailing groups into a keyboard-accessible More menu. Host
  `ToolbarButton`s get an entry there; wrap other content in `ToolbarOverflow`.
  A group holding content without an entry stays in the row, which scrolls
  horizontally when that content does not fit.

The prop-based `EditorToolbar`, `Toolbar` and `useEditorToolbar` API keeps
working unchanged in the default legacy mode and is deprecated for new
toolbars; in command mode `useEditorToolbar` returns a projection of the
command state whose callbacks run commands, and passing those props throws.

## AI agents

`onReady` hands you the open `WorkbookHandle`. An agent stages edits with
`propose()` instead of applying them; the editor paints per-cell ghosts and a
review panel where the human accepts or rejects. The full proposal API lives in
[`@betteroffice/xlsx`](https://www.npmjs.com/package/@betteroffice/xlsx).

```tsx
import { isProposalsAvailable } from "@betteroffice/xlsx";
import type { XlsxEditorApi } from "@betteroffice/xlsx-react";

<XlsxEditor
  file={file}
  onReady={({ handle, refreshProposals }: XlsxEditorApi) => {
    if (!isProposalsAvailable()) return;
    handle.propose("copilot", "add totals", [
      { sheet: 0, row: 9, col: 2, input: "=SUM(C1:C9)" },
    ]);
    refreshProposals();
  }}
/>;
```

## Edit batches

The `onReady` API's `version`, `readCells`, `findText`, `validateEdits` and
`applyEdits` run in the same queue as commands: cell and formula entries, chart
nudges, pastes and cuts accepted before them land first, and an IME composition
ends with its text written. When that cannot happen they reject with
`XlsxCommandAdmissionError`, whose `code` is `input-failed` while a refused entry
waits for correction, `gesture-active` during a chart drag, and
`document-replaced` when the workbook is replaced meanwhile. Refusals of the
batch itself stay data. `applyEdits` keeps the
caller's `expectVersion`, so input that lands first refuses the batch with
`stale-version`; read the version through the API to include it. While
`readOnly`, writes refuse with `read-only`. An applied batch repaints once and
calls `onChange` once.

```tsx
import { XlsxCommandAdmissionError, XlsxEditor, type XlsxEditorApi } from "@betteroffice/xlsx-react";

<XlsxEditor
  file={file}
  onReady={(api: XlsxEditorApi) => {
    void (async () => {
      const version = await api.version();
      const result = await api.applyEdits({
        expectVersion: version,
        steps: [
          {
            op: "setCellInputs",
            target: { sheetId: "sheet:0", range: { kind: "a1", a1: "B3" } },
            inputs: [["120"]],
          },
        ],
      });
      if (!result.ok) console.warn(result.failure.code);
    })().catch((error: unknown) => {
      if (error instanceof XlsxCommandAdmissionError) console.warn(error.code);
      else throw error;
    });
  }}
/>;
```

See [`@betteroffice/xlsx`](https://www.npmjs.com/package/@betteroffice/xlsx) for
the step vocabulary and its limits.

## Host plugins

The plugin API is experimental and may change in minor releases.

Host-owned tools (review aids, checks, templates) install through the `plugins`
prop. A plugin contributes a docked panel, an overlay on the grid and commands,
and works through restricted clients rather than the editor API or the workbook
handle:

```tsx
import { XlsxEditor, defineXlsxPlugin } from "@betteroffice/xlsx-react";

type State = { version: string | null; sheets: number };

const review = defineXlsxPlugin<State>({
  id: "acme.review",
  createState: () => ({ version: null, sheets: 0 }),
  async onEvent(context, event) {
    if (event.type !== "load" && event.type !== "document-change") return;
    const read = await context.read.readCells({ ranges: [] });
    if (read.ok) context.setState({ version: read.version, sheets: read.sheets.length }, read.version);
  },
  panel: {
    title: "Review",
    placement: "right",
    render: ({ context }) => <p>{context.state.sheets} sheets</p>,
  },
  overlay: ({ context, geometry }) => {
    const selection = context.snapshot.selection;
    const focus = selection?.cells?.focus;
    const box = selection && focus ? geometry.getCellRect({ sheetId: selection.sheetId, ...focus }) : null;
    return box ? (
      <div style={{ position: "absolute", left: box.x, top: box.y, width: box.width, height: box.height, outline: "2px dashed #2563eb" }} />
    ) : null;
  },
});

<XlsxEditor
  file={bytes}
  plugins={[review]}
  pluginGrants={{ "acme.review": { document: "write", editBatches: true } }}
  onPluginError={(error) => report(error)}
/>;
```

- **Lifecycle.** Once a workbook is open, each plugin gets fresh state,
  `initialize`, then one `load` event (`loaded`, `replaced`, or `attached` for a
  plugin added to an open workbook), and its contributions appear. A workbook
  change the hook did not make itself aborts it and delivers `load` again; after
  ten such runs the plugin is stopped and reported. It then receives
  `document-change` (the committed version after recalculation, for typing,
  pastes, commands, batches, its own included, undo, redo and remote updates,
  never for refusals or no-ops), `selection-change`, `mode-change` (with
  `readOnly`), `layout-change` and `grants-change`. Events describe current
  state: several changes may arrive as one, and a newer one aborts the hook
  still handling the previous (`context.signal`), except a change that hook's
  own edit batch made. Replacing the workbook, removing the plugin, changing its `revision`, unmounting, or a
  failure ends the activation: its signals abort, its clients refuse, and every
  `onCleanup` disposer runs once with the reason. Plugins are matched by `id`
  and `revision`, so new array or callback identities and reordering keep their
  state. `defineXlsxPlugin` copies the panel, commands and toolbar, so changing
  them takes a new definition. `context.run(action)` gives event handlers a
  fresh context and isolates their failures. `onReady` is unaffected by plugins.
- **Stale results.** `setState` returns false once the context is superseded or
  ended, or when the workbook is no longer at `atVersion` (by default the
  version the context was created at).
- **Reads and edits.** `context.read` offers `version`, `readCells`, `findText`
  and `validateEdits`, and `context.edits.applyEdits` the editor's
  version-checked batches. Each call waits for pending input in order with it,
  but plugin handlers run outside that queue, so a command can await its own
  batch.
- **Grants.** Without a grant a plugin reads, validates and navigates only, and
  `context.edits` is null. Built-in commands need their id in `commands`, and
  mutating ones also `document: "write"`; edit batches need `document: "write"`
  and `editBatches`, and `history: "none"` also `untrackedHistory`. The grant
  and `readOnly` are checked again right before each change, so a revoked grant,
  `readOnly` or a replaced workbook refuses even through a client obtained
  earlier. Mutating built-in commands have no authoritative policy yet and
  refuse plugins with `unsupported-policy`; plugins change workbooks through
  edit batches. Grants are not spreadsheet protection.
- **Contributed commands** register as `plugin:<pluginId>/<id>` on
  `api.commands`. They always run with their own plugin's clients, even when the
  toolbar, a shortcut or the host invokes them. `execute` returns
  `{ ok: true, status }` or a failure with the plugin's own code, or a refused
  edit batch as-is; callers receive it unchanged. `mutatesDocument` disables a
  command while read-only but grants nothing. Plugin shortcuts must use Mod or
  Alt, or a function key; built-in shortcuts win and a clashing plugin shortcut
  is reported. `toolbar` lists local ids in
  order: the default toolbar shows them, and replacement chrome places
  `<XlsxPluginToolbar />` where it wants them. `ToolbarCommandButton`,
  `ToolbarCommand`, `useXlsxCommand` and `useXlsxCommandState` also take a
  contributed id.
- **Selection and navigation.** `snapshot.selection` names the active sheet
  (`sheetId` and the zero-based `sheetIndex`), the selected `cells` with their
  `anchor` and `focus` direction, and the `chartId` of a selected chart.
  `navigation.selectCells({ sheetId, selection }, { expectVersion, focus })`
  activates the sheet, selects the cells and reveals the focus cell.
  `navigation.scrollToCell({ sheetId, row, col }, { expectVersion, align })`
  reveals a cell (`nearest` by default, `start` or `center`) and keeps the
  selection; a cell on another sheet activates it with nothing selected. Both
  resolve sheet ids after pending input against `expectVersion`, keep keyboard
  focus unless `selectCells` gets `focus: true`, and refuse with
  `stale-version` or `missing-target` rather than retarget. Sheet ids and
  versions are session-scoped: they do not survive saving and reopening.
- **Geometry.** `context.geometry` exists only while the canvas shows a painted
  frame of the current version. `layout` carries the frame's `sheetId`, `zoom`
  and painted `viewport` (unzoomed sheet pixels), and gets a new `id` with every
  scroll, zoom, resize and sheet switch. `getCellRect({ sheetId, row, col })`
  and `getRangeRect({ sheetId, range })` return pixels of the overlay layer,
  zoomed and clipped to the visible grid with frozen panes placed as painted;
  cells scrolled out of view or on another sheet return null. `getCellRect` is
  one grid cell; address a merged area as a range. `getPositionAtPoint` is null
  until the editor exposes pointer queries. Every method returns null once its
  layout is gone. The overlay layer sits on the grid below the editor's
  selection and cell editor, and ignores the pointer unless an element sets
  `pointer-events: auto`.
- **Panels** dock left, right or bottom of the grid, below the toolbar and
  above the sheet tabs, with tabs when several share a side. `preferredSize` is
  clamped to 40% of the workspace. In a narrow editor side docks show their
  tabs, and a tab opens its panel as a drawer that Escape closes. Collapsing a
  panel keeps the plugin running. The proposals panel opens inside the grid
  area.
- **Failures.** Every contribution renders behind its own error boundary with a
  command context restricted to its plugin. Keys pressed in plugin chrome are
  the user's: built-in shortcuts act on the editor there, a plugin's text fields
  keep their own editing keys, and no key or click in a contribution reaches the
  grid. A throwing hook, renderer, command-state function, command or
  `context.run` action stops only that plugin, runs its cleanups and reports
  `{ pluginId, generation, phase, error }` to `onPluginError`. Plugins run in
  the page's realm: grants govern the supported API, not a sandbox.

## Collaboration

Pass `collaboration` to open a network-ready replica, then attach a transport
provider from `onReady`:

```tsx
import { CollaborationProvider } from "@betteroffice/xlsx/collaboration";

<XlsxEditor
  file={file}
  collaboration={{ clientId }}
  onReady={({ handle }) => {
    const provider = new CollaborationProvider(handle, transport, {
      user: { name: "Ada" }, // shown on this peer's selection flag
    });
    provider.connect();
    return () => provider.destroy();
  }}
/>;
```

[JavaScript guide](https://docs.betteroffice.dev/docs/javascript) ·
[Changelog](https://github.com/openooxml/betteroffice/blob/main/packages/xlsx-react/CHANGELOG.md) · Apache-2.0.
