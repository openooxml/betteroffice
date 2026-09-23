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

## Host editing controls

`onSaveRequest` runs for toolbar and Ctrl/Cmd+S saves before serialization.
Return `true` to continue built-in saving; `false` or `void` handles or cancels
it. Promises are awaited and concurrent requests are coalesced. `onSave` still
receives the resulting bytes when built-in saving continues. A request waiting
on a replaced or closed document is discarded.

The API received by `onReady` exposes `flushPendingInput(): Promise<void>`.
Await it before inspecting or mutating the core from a host workflow, then call
`api.save()` for explicit serialization without re-entering `onSaveRequest`.
`save()` remains synchronous and rejects while asynchronous input is pending.
Flush rejects stale document handles, failed input, and unfinished pointer
gestures. Finish or cancel the gesture before retrying.

XLSX flushing commits cell/formula drafts and chart nudges, waits for IME
composition to finish, and waits for accepted asynchronous clipboard edits.
Reported asynchronous input failures continue to reject flushing until the
workbook is reopened.

`api.getPositionAtPoint(clientX, clientY)` returns `{ sheet, row, col }`, all
zero-based, from the painted grid. It respects scrolling and zoom without
changing focus or selection. Chart overlays, headers, outside points, stale
handles, and unavailable geometry return `null`.

For grouped host edits, use the core's existing `editCells(sheet, edits)` or
`applyOps(ops)` batch APIs. Each successful batch is one undo step. XLSX does
not currently expose editable comments, so it has no comment reanchoring API.

## What works today

- Cell editing with formula recalculation of dependents on every edit
- Editing toolbar: number formats, fonts, colors, borders, alignment, merges
- Agent proposals: in-cell tracked-change ghosts plus an accept/reject panel
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
