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

Props: `file`, `fileName`, `onSave`, `onReady` (a handle for host/agent-driven
edits), `collaboration`, `i18n`, and `className`.

## What works today

- Cell editing with formula recalculation of dependents on every edit
- Editing toolbar: number formats, fonts, colors, borders, alignment, merges
- Agent proposals: in-cell tracked-change ghosts plus an accept/reject panel
- Version-checked edit batches through `onReady`, after pending input commits
- TSV clipboard copy/paste
- Accessible grid mirroring the painted canvas for screen readers
- Localized UI via the `i18n` prop
  ([`@betteroffice/xlsx-i18n`](https://www.npmjs.com/package/@betteroffice/xlsx-i18n))
- Real-time collaboration with people or agents; the workbook is a CRDT
- Live collaborator selections and presence chips, shown in each peer's color

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
`applyEdits` first commit pending input — cell and formula drafts, chart nudges
and clipboard pastes in flight — and reject while text composition or a chart
drag is unfinished or the workbook is replaced meanwhile. `applyEdits` keeps the
caller's `expectVersion`, so input that lands first refuses the batch with
`stale-version`; read the version through the API to include it. While
`readOnly`, writes refuse with `read-only`. An applied batch repaints once and
calls `onChange` once.

```tsx
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
    })();
  }}
/>;
```

See [`@betteroffice/xlsx`](https://www.npmjs.com/package/@betteroffice/xlsx) for
the step vocabulary and its limits.

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
