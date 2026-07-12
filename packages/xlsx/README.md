# @betteroffice/xlsx

Framework-free core for the BetterOffice XLSX editor — the Rust engine (parse,
calc, render) compiled to WebAssembly and bundled in, plus display-list,
viewport, hit-test, and accessibility helpers. There's no separate wasm asset to
host.

> **Experimental (`0.0.x`).** The API is unstable and may change in any release.

```bash
bun add @betteroffice/xlsx
```

Most apps want the turnkey React component in
[`@betteroffice/xlsx-react`](https://www.npmjs.com/package/@betteroffice/xlsx-react).
Reach for this package directly to render a spreadsheet onto your own canvas.

## Usage

```ts
import { openWorkbook } from "@betteroffice/xlsx";

// parse + load into the wasm engine (calc-ready)
const bytes = new Uint8Array(await file.arrayBuffer());
const workbook = openWorkbook(bytes);
```

From the returned handle you compose the geometry and rendering helpers this
package exports onto your own `<canvas>` 2D context — `paintDisplayList`,
`cellAtPoint` / `cellRect` / `rangeRect` (hit-testing), the `viewport` math,
`buildA11yGrid`, and `toTsv` / `fromTsv` (clipboard).

Docs: https://betteroffice.dev · Apache-2.0.
