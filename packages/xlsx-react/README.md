# @betteroffice/xlsx-react

React chrome for the BetterOffice XLSX editor — wraps
[`@betteroffice/xlsx`](https://www.npmjs.com/package/@betteroffice/xlsx) in a
drop-in `<XlsxEditor>` component with selection, keyboard, and clipboard wired up.

> **Experimental (`0.0.x`).** The API is unstable and may change in any release.

```bash
bun add @betteroffice/xlsx-react @betteroffice/xlsx react react-dom
```

`react` and `react-dom` (18 or 19) are peer dependencies.

## Usage

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
        onSave={(bytes) => {
          // edited .xlsx bytes; without onSave the save button downloads them
          console.log(`saved ${bytes.length} bytes`);
        }}
      />
    </>
  );
}
```

Props: `file` (a `Uint8Array` of `.xlsx` bytes — omit it to render an empty
frame), `fileName`, `onSave`, `onReady` (a handle for host/agent-driven edits),
and `className`.

Docs: https://betteroffice.dev · Apache-2.0.
