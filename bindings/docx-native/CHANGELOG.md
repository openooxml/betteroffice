# @betteroffice/docx-native

## 0.6.0

### Minor Changes

- c9a8850: The native `EditSession` now also lays out, builds display frames, measures fonts, takes resident input, hit tests and reports update callbacks, and every WebAssembly engine function is exported with the same name and results.
- e9fbbe6: Replaces the native facade with the same entry points and API as `@betteroffice/docx`, backed by the native engine by default. Import `@betteroffice/docx-native` for Node.js, Bun and Deno without WebAssembly initialization or assets.

### Patch Changes

- e1a414e: Documents whose XML parts start with a byte order mark and have bookmarks between paragraphs now save. Before, saving them failed with a malformed XML error.
- 3c1c056: Native DOCX page rendering is up to ten times faster on pages with clipped content. Pixel output is unchanged.
- 0a4285c: Adds opt-in `fontPreflight`, returning DOCX font faces, script keys, and scan completeness without the editing engine. Its standalone module loads only when called or explicitly loaded with `preloadPreflightWasm`.
- 773c633: Keeps bookmarks in table cells, content controls, headers, footers, notes, comments and text boxes on save, including after blocks are deleted, merged or moved.
- 2dbba20: `@betteroffice/docx-native` installs the Latin, international and CJK font packs and reads each default face from disk on first use, so international and CJK text gets its bundled face with nothing fetched.
- 514e876: Story checksums use constant memory, so long documents no longer exhaust the WebAssembly engine's memory and stop the editor. Checksum values are unchanged.
- ad45f86: Adds opt-in `decodeFontFace` for packed fonts. Format package roots return a promise of `font` and optional `payload` ArrayBuffers; native bindings return byte arrays synchronously. Invalid options and decode errors reject or throw without retrying another decoder.
- ec04baa: The native packages require Node.js 20.16+ or 22.3+, which they need to read installed font packs from disk.
- Updated dependencies [de8f092]
- Updated dependencies [5999962]
- Updated dependencies [9ab2b61]
- Updated dependencies [39ae19a]
- Updated dependencies [ec04baa]
- Updated dependencies [df33f6a]
- Updated dependencies [df33f6a]
- Updated dependencies [37199d9]
- Updated dependencies [842d624]
- Updated dependencies [e65df1c]
- Updated dependencies [ad45f86]
  - @betteroffice/fonts-cjk@1.0.0
  - @betteroffice/fonts@1.0.0
  - @betteroffice/fonts-intl@1.0.0
  - @betteroffice/fonts-latin@1.0.0

## 0.5.1

## 0.5.0

### Patch Changes

- c688743: Adds prebuilt Rust engines for Node.js, Bun and Deno that open, edit, render PNGs and save documents. Image decoding stays within the render budget.
