# @betteroffice/xlsx-native

## 0.6.0

### Patch Changes

- ad45f86: Adds opt-in `decodeFontFace` for packed fonts. Format package roots return a promise of `font` and optional `payload` ArrayBuffers; native bindings return byte arrays synchronously. Invalid options and decode errors reject or throw without retrying another decoder.
- ec04baa: The native packages require Node.js 20.16+ or 22.3+, which they need to read installed font packs from disk.
- 4990e8d: Native XLSX rendering draws charts much faster: a chart sheet that took about 400 ms now takes about 20 ms. Pixel output is unchanged.
- 7a91087: Now exposes exactly the `@betteroffice/xlsx` API (same entry points, exports and types) on the native engine for Node.js, Bun and Deno; pick the engine by import. The `browser` condition throws. The `openWorkbook(Buffer)`/`Workbook` facade is removed.

## 0.5.1

## 0.5.0

### Patch Changes

- c688743: Adds prebuilt Rust engines for Node.js, Bun and Deno that open, edit, render PNGs and save documents. Image decoding stays within the render budget.
