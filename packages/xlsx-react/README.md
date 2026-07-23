# @betteroffice/xlsx-react

React chrome for the BetterOffice XLSX editor — wraps
[`@betteroffice/xlsx`](https://www.npmjs.com/package/@betteroffice/xlsx) in a
drop-in `<XlsxEditor>` component with the toolbar, selection, keyboard, and
clipboard wired up. Formula calculation and rendering run in the core's
Rust/WebAssembly engine; the grid is painted onto canvas.

<!-- TODO(author): screenshot/GIF — editor with agent proposal ghosts / collab cursors -->
<img src="https://betteroffice.dev/readme/xlsx-editor.png" alt="XlsxEditor with the editing toolbar and a pending agent proposal" width="720" />

> **Early (`0.0.x`).** The core surfaces — opening/saving documents, the editor
> components, collaboration — are settling and unlikely to change shape. Smaller
> APIs may still move between releases; breaking changes are always listed in
> the changelog.

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
- TSV clipboard copy/paste
- Accessible grid mirroring the painted canvas for screen readers
- Localized UI via the `i18n` prop
  ([`@betteroffice/xlsx-i18n`](https://www.npmjs.com/package/@betteroffice/xlsx-i18n))
- Real-time collaboration with people or agents; the workbook is a CRDT
- Live collaborator cursors are landing in the next release.

## AI agents

`onReady` hands you the open `WorkbookHandle`. An agent stages edits with
`propose()` instead of applying them; the editor paints per-cell ghosts and a
review panel where the human accepts or rejects. The full proposal API lives in
[`@betteroffice/xlsx`](https://www.npmjs.com/package/@betteroffice/xlsx).

```tsx
import type { XlsxEditorApi } from "@betteroffice/xlsx-react";

<XlsxEditor
  file={file}
  onReady={({ handle, refreshProposals }: XlsxEditorApi) => {
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
    const provider = new CollaborationProvider(handle, transport);
    provider.connect();
    return () => provider.destroy();
  }}
/>;
```

Docs: https://betteroffice.dev · Apache-2.0.
