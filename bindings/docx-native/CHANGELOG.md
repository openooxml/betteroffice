# @betteroffice/docx-native

## 0.6.0

### Minor Changes

- c9a8850: The native `EditSession` now also lays out, builds display frames, measures fonts, takes resident input, hit tests and reports update callbacks, and every WebAssembly engine function is exported with the same name and results.
- 905ef77: The raw `EditSession.renderPngTarget(displayListJson, targetJson, optionsJson)` rasterizes one laid-out page to PNG in the native engine, from the same display list the canvas paints, with the session's fonts. Oversized pages and missing fonts are refused with stable messages.
- e9fbbe6: Replaces the native facade with the same entry points and API as `@betteroffice/docx`, backed by the native engine by default. Import `@betteroffice/docx-native` for Node.js, Bun and Deno without WebAssembly initialization or assets.
- e20cdb3: `YrsSession.renderPng({ page }, options?)` draws a page as PNG bytes with the editor's layout and painter, or natively with `@betteroffice/docx-native`. It rejects with `PngExportError` rather than paint a substitute font.
- b2ea7f5: The docx and pptx native bindings use mimalloc as their memory allocator, through the default-on `mimalloc` cargo feature. DOCX save, edit and layout get faster and PPTX open and export faster, with a higher peak memory on some documents.

### Patch Changes

- 1282f77: Adds opt-in `set_language` edits for DOCX paragraphs, grep matches, styles and document defaults. Latin, East Asian and complex-script proofing language slots can be set or cleared independently while preserving other formatting.
- c0ae1af: Formatting a paragraph mark inside a block content control, such as a form field, no longer rewrites the rest of the document on save: tracked changes, proofing marks, drawings and fields elsewhere are kept as they were.
- e1a414e: Documents whose XML parts start with a byte order mark and have bookmarks between paragraphs now save. Before, saving them failed with a malformed XML error.
- c2b9ac7: Tables inserted or pasted in the editor lay out with the cell margins of the document's table styles, as Word shows them once saved, so text wraps on screen where Word wraps it. Opened tables are unchanged.
- 3c1c056: Native DOCX page rendering is up to ten times faster on pages with clipped content. Pixel output is unchanged.
- a9b74ab: Preserves bookmarks inside complex field results when saving, including fields that span paragraphs in the document body or table cells.
- 0a4285c: Adds opt-in `fontPreflight`, returning DOCX font faces, script keys, and scan completeness without the editing engine. Its standalone module loads only when called or explicitly loaded with `preloadPreflightWasm`.
- c3c5294: Picture watermarks that Word also places in a footer are kept when the document is saved.
- c2281b5: Formatting a selection that spans a table, such as select all then bold, now formats the text in its cells. Paragraphs a save would lose content from stay unchanged, and edits and saves report which and why.
- fa2717d: Every engine now reads decimal numbers back exactly. Drawing groups whose scaled child coordinates have long decimals stay intact on save.
- 773c633: Keeps bookmarks in table cells, content controls, headers, footers, notes, comments and text boxes on save, including after blocks are deleted, merged or moved.
- 8efeef9: Keeps bookmarks inside field codes, such as the one Word writes for each legacy form field, and fields nested inside another field's code, on save.
- 51a1bbf: Keeps drawing groups and canvases on save in headers, footers and rewritten paragraphs. A group that matches its source apart from its frame and text-box text is written back from the source with those changes; any other edit writes it from the model as before.
- 30134ef: Keeps bookmarks that sit between table rows or between cells on save, including after rows or columns are inserted or deleted in an editing session.
- bf8880c: Keeps bookmarks inside the paragraphs of text boxes on save, including the copies Word stores for older versions.
- f777004: Keeps a text box on its own run on save when bookmarks, comment ranges or other markers come before it in the paragraph; such a text box was dropped or moved to another run.
- f31b62e: Keeps text boxes inside tracked insertions, deletions, hyperlinks and simple fields on save, with their text and revision marks.
- dbcc3fd: Undo grouping reads JavaScript `Date.now`, as the wasm engine does, so edits group the same way on both engines, including under a mocked clock.
- 2dbba20: `@betteroffice/docx-native` installs the Latin, international and CJK font packs and reads each default face from disk on first use, so international and CJK text gets its bundled face with nothing fetched.
- 56c2c6c: Native docx page PNGs draw SVG, EMF and WMF pictures. A picture the renderer still cannot draw rejects `renderPng` with the new `picture-unavailable` code instead of leaving it blank.
- 934104f: A table pasted from another document lays out with this document's Word compatibility mode, as Word shows it once saved, instead of the source document's, so it no longer shifts left by its cell margin.
- 22fc562: Clearing or changing character spacing, scaling, kerning or raised text, and clearing a theme font, no longer rebuilds the paragraph on save, so its hyperlinks, smart tags and drawings are kept as they were.
- 6d54917: Native page PNGs clip elliptical pictures and their borders to the ellipse, as the editor does, instead of drawing a rectangle.
- a2d41a8: Native page PNGs draw picture-filled shapes, stretched or tiled, with their outlines, as the editor does, instead of leaving the shape out.
- aa54bb8: Native page PNGs accept VML color names, and a list marker in the paragraph's default font no longer refuses the page with a missing font chain.
- cfdeb9e: Saving a document after an edit that changes many paragraphs, such as clearing formatting over the whole body, is much faster on large documents. Saved files are unchanged.
- 514e876: Story checksums use constant memory, so long documents no longer exhaust the WebAssembly engine's memory and stop the editor. Checksum values are unchanged.
- 3056f24: Table row borders can be dragged to set a row's height in both docx editors. The height saves as an at-least `w:trHeight`, and a drag is one undo step. Adds `set_row_height` to the engine.
- c74b2b8: The docx editing engine is smaller to download.
- c2e450c: Text boxes take the document's default East Asian language, so Japanese and Korean text boxes draw with their own CJK face. A font used by several CJK scripts puts the document's own script first.
- c87ef90: Text boxes apply all package paragraph and character styles by default, including inherited fonts, colors, languages and spacing. Shared state preserves these styles across peers and payload edits without changing saved payloads. `font_preflight` includes styled text-box faces and inherited language scripts.
- 0bee43d: Text runs drawn without shaping carry the script faces their paragraph needs (`scripts`), so the native renderer requests the Japanese, Korean or Traditional Chinese face the browser path already chooses.
- ad45f86: Adds opt-in `decodeFontFace` for packed fonts. Format package roots return a promise of `font` and optional `payload` ArrayBuffers; native bindings return byte arrays synchronously. Invalid options and decode errors reject or throw without retrying another decoder.
- 1af9969: Adds prebuilt binaries for musl Linux on x64 and arm64, such as Alpine, and for Windows on arm64.
- ec04baa: The native packages require Node.js 20.16+ or 22.3+, which they need to read installed font packs from disk.
- 681b02e: A Rust panic in a native engine now fails its session as a trap, as on WebAssembly, instead of aborting the process. The engine object that panicked refuses every later call except free.
- 6378188: Worker sessions now run on Node.js >= 20.16 / 22.3 through `worker_threads`, keeping the process alive until disposed. DOCX worker editing needs `OffscreenCanvas` and fails with a clear error without it (Node.js, Bun, Deno); open, read and save still work.
- 7cf70f8: Worker entries no longer import the module that constructs their own worker, so Next.js (Turbopack) production builds of apps that use `./worker` finish instead of hanging.
- 55c25ba: Worker files and main-thread entries import no private `#` specifiers, so worker sessions open from a CDN or with an import map, without a bundler. On Node.js, a worker thread that fails to start now fails the session instead of throwing synchronously.
- 6d79fc5: Undo tracking no longer slows with the number of stories, so large deletions and clears in documents with many table cells stay fast. Published Rust crates get this once yrs ships it.
- Updated dependencies [f99e53a]
- Updated dependencies [de8f092]
- Updated dependencies [3609ce4]
- Updated dependencies [5999962]
- Updated dependencies [d60b8e4]
- Updated dependencies [9ab2b61]
- Updated dependencies [39ae19a]
- Updated dependencies [d60b8e4]
- Updated dependencies [53502c6]
- Updated dependencies [ec04baa]
- Updated dependencies [d60b8e4]
- Updated dependencies [21a60e5]
- Updated dependencies [19c20c3]
- Updated dependencies [df33f6a]
- Updated dependencies [df33f6a]
- Updated dependencies [37199d9]
- Updated dependencies [842d624]
- Updated dependencies [e65df1c]
- Updated dependencies [18b3628]
- Updated dependencies [dc78510]
- Updated dependencies [ad45f86]
- Updated dependencies [e6d5c52]
  - @betteroffice/fonts@1.0.0
  - @betteroffice/fonts-cjk@1.0.0
  - @betteroffice/fonts-intl@1.0.0
  - @betteroffice/fonts-latin@1.0.0

## 0.5.1

## 0.5.0

### Patch Changes

- c688743: Adds prebuilt Rust engines for Node.js, Bun and Deno that open, edit, render PNGs and save documents. Image decoding stays within the render budget.
