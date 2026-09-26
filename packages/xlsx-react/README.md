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
  during a chart drag, or `proposal-stale`. A cell entry the workbook refuses
  stays open at its own cell for correction; until it is corrected or discarded
  with Escape, every command fails with `input-failed`. Print waits until the
  canvas has painted the written text, and fails with `render-failed` when it
  cannot.
- `api.save()` stays synchronous and never returns bytes without accepted
  input: it throws `XlsxSaveRefusedError` with `code` `input-pending` while
  entries or a paste still wait to be written, and `input-failed` while a
  refused entry waits for correction. `await api.commands.execute("save", null)`
  waits for that input and then saves through `onSave`.
- Every cell write goes through one queue. `api.selectCells`, `api.clearSelection`
  and switching sheets close the open entry and write it at once when nothing
  waits; otherwise it is queued behind the earlier input, so it may not have
  landed when they return. An open entry the workbook refuses stays open and
  they change nothing (`selectCells` returns `false`); one it refused earlier
  reopens at its own cell once they land on its sheet. Commands run right
  after them, even in the same handler, act on the new selection and sheet.
- The store enforces read-only mode, the selection a command needs and its
  arguments for the editor's UI. `api.handle` stays unrestricted host
  authority, and the store is not workbook protection.
- Narrow toolbars move trailing groups into a keyboard-accessible More menu. Host
  `ToolbarButton`s get an entry there; wrap other content in `ToolbarOverflow`.
  A group holding content without an entry stays in the row.

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
