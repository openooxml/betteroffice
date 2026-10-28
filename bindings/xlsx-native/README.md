# @betteroffice/xlsx-native

The native engine for `@betteroffice/xlsx` on Node.js, Bun and Deno. Install it
next to the matching version of `@betteroffice/xlsx`:

```sh
npm install @betteroffice/xlsx @betteroffice/xlsx-native
```

`@betteroffice/xlsx` picks this engine automatically when a matching native
binary is available, and falls back to WebAssembly otherwise. Prebuilt binaries
cover macOS arm64/x64, GNU Linux arm64/x64, and Windows x64.

```js
import { engineBackend } from '@betteroffice/xlsx';

console.log(engineBackend());
```

`engineBackend()` reports `'native'` or `'wasm'`. Set `BETTEROFFICE_ENGINE=wasm`
to force WebAssembly, or `BETTEROFFICE_ENGINE=native` to require the native
engine. Requiring native throws if the binary cannot load or its version does
not match `@betteroffice/xlsx`.

## Deprecated facade

The asynchronous `openWorkbook` and `Workbook` facade exported by this package
is deprecated and removed in 0.7. Use `@betteroffice/xlsx`, which runs on this
native engine when it is installed. Initialize the core once with `initWasm`
before opening workbooks; pass WebAssembly bytes for the fallback in server
runtimes. Edits and saves are synchronous. Cell row and column indexes are
zero-based, so A1 is row 0, column 0.

| 0.5 facade | @betteroffice/xlsx |
| --- | --- |
| `await openWorkbook(buf, { recalculate, nowSerial })` | `openWorkbook(bytes, { calculation: { nowSerial, randSeed } })` |
| `await wb.set(0, 'A1', v)` | `wb.editCell(0, row, col, v)` |
| `await wb.setStyle(0, 'A1:B1', patch)` | `wb.patchRangeStyle(0, 'A1:B1', patch)` |
| `await wb.save()` (Buffer) | `wb.save()` (Uint8Array) |
