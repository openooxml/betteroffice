# @betteroffice/xlsx-native

Native Rust bindings for the BetterOffice XLSX engine in Node.js, Bun and Deno.
Prebuilt binaries cover macOS arm64/x64, GNU Linux arm64/x64, and Windows x64.

```sh
npm install @betteroffice/xlsx-native
```

## Deprecated facade

The asynchronous `openWorkbook` and `Workbook` facade exported by this package
is deprecated and removed in 0.7. Use `@betteroffice/xlsx` instead. Initialize
the core once with `initWasm` before opening workbooks; pass WebAssembly bytes
in server runtimes. Edits and saves are synchronous. Cell row and column indexes are
zero-based, so A1 is row 0, column 0.

| 0.5 facade | @betteroffice/xlsx |
| --- | --- |
| `await openWorkbook(buf, { recalculate, nowSerial })` | `openWorkbook(bytes, { calculation: { nowSerial, randSeed } })` |
| `await wb.set(0, 'A1', v)` | `wb.editCell(0, row, col, v)` |
| `await wb.setStyle(0, 'A1:B1', patch)` | `wb.patchRangeStyle(0, 'A1:B1', patch)` |
| `await wb.save()` (Buffer) | `wb.save()` (Uint8Array) |

The facade keeps the cached formula values unless `recalculate` is true;
`@betteroffice/xlsx` always recalculates on open, so formula values that were
stale in the file can change. A `calculation` context needs both `nowSerial`
and `randSeed`; without one, `NOW()` and `TODAY()` use the current time.
