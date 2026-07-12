# @betteroffice/xlsx

Framework-free core for the BetterOffice XLSX editor — the Rust engine (parse,
calc, render) compiled to WebAssembly, plus display-list, viewport, hit-test, and
accessibility helpers.

> **Experimental (`0.0.x`).** The API is unstable and may change in any release.

```bash
bun add @betteroffice/xlsx
```

Most apps want the turnkey React component in
[`@betteroffice/xlsx-react`](https://www.npmjs.com/package/@betteroffice/xlsx-react).
Reach for this package directly to render a spreadsheet onto your own canvas.

## Usage

```ts
import { initWasm, openWorkbook } from "@betteroffice/xlsx";

await initWasm();
const bytes = new Uint8Array(await file.arrayBuffer());
const workbook = openWorkbook(bytes);
```

`initWasm()` fetches the packaged wasm asset once in browsers. Pass wasm bytes or
a precompiled `WebAssembly.Module` explicitly in runtimes that cannot fetch the
asset URL.

## Development

The generated `.wasm` binary is intentionally not committed. From the repository
root, install `wasm-pack` 0.15.0 and run `bun run build:xlsx-wasm`. Package builds,
tests, demo startup, and CI run this step automatically.

From the returned handle you compose the geometry and rendering helpers this
package exports onto your own `<canvas>` 2D context — `paintDisplayList`,
`cellAtPoint` / `cellRect` / `rangeRect` (hit-testing), the `viewport` math,
`buildA11yGrid`, and `toTsv` / `fromTsv` (clipboard).

Docs: https://betteroffice.dev · Apache-2.0.
