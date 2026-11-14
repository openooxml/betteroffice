# @betteroffice/docx-native

## 0.6.0

### Minor Changes

- 78f28c5: `applyEdits` text steps and new paragraphs accept `runs` with bold, italic, underline, strike, colour, highlight, size, font and superscript/subscript marks instead of plain `text`. Rust `EditOperation` and `ParagraphInput` gain a `runs` field.
- 2a1d763: Adds the `formatText` edit step (`set`, `clear` back to the styles, `onUnsupported`). The agents `format_text` tool gains strike, superscript/subscript, colour, highlight, size and font.
- 57fa9d7: Screen reader page mirrors read a visible list bullet as a localised word, never its stored character. `DocAttrs` gains `listMarkerBullet` and `MirrorLabels` gains `bullet`; without either, a bullet stays hidden.
- b1a4b01: Screen reader page mirrors read list numbers and labels such as "3.2" or "Appendix 1", which were hidden. Bullets stay hidden.
- c9a8850: The native `EditSession` now also lays out, builds display frames, measures fonts, takes resident input, hit tests and reports update callbacks, and every WebAssembly engine function is exported with the same name and results.
- 590615b: The native DOCX engine can report its editing heap per thread. Counting is off by default; `startHeapCounting()` from `@betteroffice/docx/yrs` turns it on. Pressure events on native need a host memory budget.
- 905ef77: The raw `EditSession.renderPngTarget(displayListJson, targetJson, optionsJson)` rasterizes one laid-out page to PNG in the native engine, from the same display list the canvas paints, with the session's fonts. Oversized pages and missing fonts are refused with stable messages.
- e9fbbe6: Replaces the native facade with the same entry points and API as `@betteroffice/docx`, backed by the native engine by default. Import `@betteroffice/docx-native` for Node.js, Bun and Deno without WebAssembly initialization or assets.
- c0e297a: The engine reads the text a laid-out page shows, built or not: `page_texts_json` returns each page's paragraphs, cut where the page cuts them, with heading levels, and the table rows the page places.
- a9897ef: DOCX PNG export tells onDiagnostic, once per page, which kinds of effect the engine does not draw (`effects-skipped`). The shared PNG diagnostic type gains that code, so a host with an exhaustive switch on the code gets a new case.
- 62605af: DOCX PNG export draws a placeholder for a picture it cannot draw instead of rejecting the page or leaving it blank; `onDiagnostic` reports each one and why. `strictPictures: true` keeps rejecting. Agent previews warn.
- e20cdb3: `YrsSession.renderPng({ page }, options?)` draws a page as PNG bytes with the editor's layout and painter, or natively with `@betteroffice/docx-native`. It rejects with `PngExportError` rather than paint a substitute font.
- ae6e7e1: Fixes collaborative documents that stopped laying out after an undo removed a table cell another collaborator's table still showed. That cell now reads as empty and refuses edits until the undo is redone or its row is deleted.
- 6d8fb19: A table that crossing undos left unlayoutable is now skipped instead of breaking layout; table commands and saves refuse with a coded error until redo restores it or it is deleted. Both engines add `unavailable_tables_from`.
- b2ea7f5: The docx and pptx native bindings use mimalloc as their memory allocator, through the default-on `mimalloc` cargo feature. DOCX save, edit and layout get faster and PPTX open and export faster, with a higher peak memory on some documents.
- f78484e: Adds `prime_last_slice_cost` to the wasm and native DOCX edit sessions: what the last `prime_save_seed_states_slice` call charged against its byte budget, overshoot included, or 0 (committed-scan slices charge nothing). The docx worker sizes the next slice from it, growing at most twofold.

### Patch Changes

- 1282f77: Adds opt-in `set_language` edits for DOCX paragraphs, grep matches, styles and document defaults. Latin, East Asian and complex-script proofing language slots can be set or cleared independently while preserving other formatting.
- 9590d1e: Paragraph-relative shapes now move with an anchor paragraph whose first line overflows to the next page, instead of staying behind on the previous page.
- 232fe11: Text set in Simplified Arabic or Traditional Arabic now uses the line height Word uses for those fonts, so Arabic documents break pages closer to where Word breaks them.
- c3d6f21: Autofit tables with no table width keep their cells' stated widths when those widths and the column grid both exceed the text width, as Word does, instead of squeezing the columns and wrapping text that fits.
- c0ae1af: Formatting a paragraph mark inside a block content control, such as a form field, no longer rewrites the rest of the document on save: tracked changes, proofing marks, drawings and fields elsewhere are kept as they were.
- e1a414e: Documents whose XML parts start with a byte order mark and have bookmarks between paragraphs now save. Before, saving them failed with a malformed XML error.
- c2b9ac7: Tables inserted or pasted in the editor lay out with the cell margins of the document's table styles, as Word shows them once saved, so text wraps on screen where Word wraps it. Opened tables are unchanged.
- 455254f: Native DOCX rendering no longer fails on pages with charts or comma-separated font names. Chart text naming a theme font now uses the document theme's font in every engine.
- fedbb82: Clear Formatting in the DOCX editor returns text to what its paragraph and table styles give, so headings stay bold. `formatText` clear `'all'` no longer changes a run's text direction. Paragraph patches gain an optional `inheritance` field.
- 3c1c056: Native DOCX page rendering is up to ten times faster on pages with clipped content. Pixel output is unchanged.
- 46c17cb: Documents whose styles declare empty document defaults lay out their unstyled paragraphs and table cells without paragraph spacing, as Word does, instead of with Word's 8 pt after and 1.08 line spacing.
- c13d981: A document whose last paragraph holds only a page break now starts its endnotes under the last line of text, as Word does, instead of on a new page.
- 33fcb87: Endnotes now follow the last body text, as in Word, and continue onto new pages under a continuation separator. Before, long endnotes pushed body text onto extra pages and were painted over it.
- a9b74ab: Preserves bookmarks inside complex field results when saving, including fields that span paragraphs in the document body or table cells.
- 0a4285c: Adds opt-in `fontPreflight`, returning DOCX font faces, script keys, and scan completeness without the editing engine. Its standalone module loads only when called or explicitly loaded with `preloadPreflightWasm`.
- c3c5294: Picture watermarks that Word also places in a footer are kept when the document is saved.
- c2281b5: Formatting a selection that spans a table, such as select all then bold, now formats the text in its cells. Paragraphs a save would lose content from stay unchanged, and edits and saves report which and why.
- fa2717d: Every engine now reads decimal numbers back exactly. Drawing groups whose scaled child coordinates have long decimals stay intact on save.
- 8e4260c: A DOCX table row marked hidden keeps its height when any paragraph mark in it is visible, as Word does, even if the paragraphs are empty. Only rows whose paragraph marks and runs are all hidden are left out.
- fc50a58: DOCX table rows marked hidden (`w:trPr/w:hidden`) whose content is all hidden no longer paint borders or shading or take height unless hidden text is shown. Marked rows with visible content still render. A fully hidden table takes no space; cells keep their document row numbers.
- b756b5f: Runs whose character scale is 0, which Word draws at full width, now draw at full width instead of disappearing on canvas and failing the page natively. Other values outside 1 to 600 are ignored or held to 600.
- 773c633: Keeps bookmarks in table cells, content controls, headers, footers, notes, comments and text boxes on save, including after blocks are deleted, merged or moved.
- 8efeef9: Keeps bookmarks inside field codes, such as the one Word writes for each legacy form field, and fields nested inside another field's code, on save.
- 51a1bbf: Keeps drawing groups and canvases on save in headers, footers and rewritten paragraphs. A group that matches its source apart from its frame and text-box text is written back from the source with those changes; any other edit writes it from the model as before.
- 30134ef: Keeps bookmarks that sit between table rows or between cells on save, including after rows or columns are inserted or deleted in an editing session.
- bf8880c: Keeps bookmarks inside the paragraphs of text boxes on save, including the copies Word stores for older versions.
- f777004: Keeps a text box on its own run on save when bookmarks, comment ranges or other markers come before it in the paragraph; such a text box was dropped or moved to another run.
- f31b62e: Keeps text boxes inside tracked insertions, deletions, hyperlinks and simple fields on save, with their text and revision marks.
- 002ac93: An open session now repaginates like a freshly opened one after edits next to a page break or to a continuous section's margins. `PageFlowGeometry` can no longer be built with a struct literal outside docx-layout.
- afcb67e: A DOCX list level override that sets no start value now keeps the start of the level it overrides, so its first label reads "A." or "3." instead of nothing or "0.".
- 10fbc38: DOCX numbered list levels without a start value now count from zero, as Word does, and a zero counter prints "0." instead of nothing.
- fa27c3c: DOCX lines measured with authoritative shaping now carry the glyphs measurement shaped them with, as `glyphRuns`, packed into one string per line. Rendering does not use them yet.
- 464d5c6: Screen reader page mirrors read a chart as a figure holding its values as a table, series as column headers and categories as row headers, formatted like its data labels. A chart without cached values says it has no data.
- 59d467a: The accessible page text reads a line break inside a paragraph as a new line and a tab as a tab, so the words on either side no longer run together for screen readers.
- 140c8fa: Screen reader page mirrors read the first column of a table with "Repeat as header row" rows and first-column formatting as row headers, and data cells name their row header. Column headers are unchanged.
- 29dc450: Text with Word's classic shadow draws it instead of an empty Word 2010 shadow, a Word 2010 shadow no longer adds a classic one and keeps its own settings, and a content control's data binding is no longer replaced by its Word 2013 copy.
- 57f4975: Arabic text set in Simplified Arabic or Traditional Arabic now falls back to the bundled Naskh face, closer to those fonts' style and width.
- 452e0a9: Native DOCX PNG export paints rectangular and path gradient fills as a radial gradient, as the browser renderer does, instead of failing the page.
- 533fb57: The native DOCX engine now enforces the worker heap limit: a call past it fails with the same out-of-memory error as wasm, and its object refuses later calls. A limit turns heap counting on; without a limit or `startHeapCounting()`, counting costs nothing.
- dbcc3fd: Undo grouping reads JavaScript `Date.now`, as the wasm engine does, so edits group the same way on both engines, including under a mocked clock.
- 2dbba20: `@betteroffice/docx-native` installs the Latin, international and CJK font packs and reads each default face from disk on first use, so international and CJK text gets its bundled face with nothing fetched.
- 4520944: `install_panic_hook` now works and the package installs it on load: on macOS and Linux a Rust panic writes one `[docx-native] panic:` line to stderr, then Rust's default report. The thrown error is unchanged; calling it again does nothing.
- 56c2c6c: Native docx page PNGs draw SVG, EMF and WMF pictures. A picture the renderer still cannot draw rejects `renderPng` with the new `picture-unavailable` code instead of leaving it blank.
- ce686a8: Footnote and endnote reference marks in the document view now print the note's number, in its format and from its starting value, as the note area shows it, instead of its internal id, also while the per-revision accept/reject preview hides or shows a reference.
- fa08d5d: Word documents no longer form ligatures such as fi unless the document asks for them, matching Word; contextual alternates follow the document too. The caret and selection land between a ligature's letters, and selecting a field result covers all of it.
- ed0a3ce: DOCX Arabic, Syriac and Devanagari now paint the glyphs measurement shaped, so words join on screen, including words split across runs that differ only outside the complex-script font.
- 934104f: A table pasted from another document lays out with this document's Word compatibility mode, as Word shows it once saved, instead of the source document's, so it no longer shifts left by its cell margin.
- 22fc562: Clearing or changing character spacing, scaling, kerning or raised text, and clearing a theme font, no longer rebuilds the paragraph on save, so its hyperlinks, smart tags and drawings are kept as they were.
- 7796969: Typing next to a field whose begin, code, result and end all sit in one paragraph no longer rebuilds that paragraph on save, so its other content is kept as it was.
- 914dabc: DOCX PNG export names a picture whose part is missing by its relationship id, in `onDiagnostic` and in the `strictPictures` rejection, instead of leaving it unnamed. Display-list pictures carry it as `relationshipId`.
- 6d54917: Native page PNGs clip elliptical pictures and their borders to the ellipse, as the editor does, instead of drawing a rectangle.
- a2d41a8: Native page PNGs draw picture-filled shapes, stretched or tiled, with their outlines, as the editor does, instead of leaving the shape out.
- aa54bb8: Native page PNGs accept VML color names, and a list marker in the paragraph's default font no longer refuses the page with a missing font chain.
- 94093ba: `parse_docx_s4` no longer fails on an unreadable relationships part other than the document's own. It reports a warning and leaves that part's references unresolved.
- 524652d: A page or column break that an edit moves to the start of the next paragraph is now saved there instead of being dropped, in paragraphs that start with a plain run.
- cfdeb9e: Saving a document after an edit that changes many paragraphs, such as clearing formatting over the whole body, is much faster on large documents. Saved files are unchanged.
- cda31aa: Saving a document records paragraph ids in linear time instead of comparing every saved id with every live paragraph. At 5 MiB this step drops from about 2.7 s to 63 ms on native and 4.9 s to 0.2 s on wasm; output is unchanged.
- a711586: When a DOCX is first loaded for editing, equal attribute keys and values share one allocation. Memory retained after open drops by 1 to 6% on large real documents (10% on a text-heavy synthetic 5 MiB file). Peak memory and saved bytes are unchanged.
- 84b271b: Opening a DOCX keeps equal embed keys and values in one shared allocation: retained memory after open falls 7.8 to 12.0% on the two largest public test documents and 29.2 to 30.5% on a 5 MiB synthetic document, native and wasm. Saved bytes are unchanged.
- 946f6f5: Picture fills in SmartArt now resolve against the diagram drawing's own relationships, so they draw instead of failing the page natively and falling back to a solid fill on canvas.
- 514e876: Story checksums use constant memory, so long documents no longer exhaust the WebAssembly engine's memory and stop the editor. Checksum values are unchanged.
- bbac3e7: Fixes collaborative table editing where splitting a cell again after a collaborator removed the previous split cell made the new cell invisible, and where redoing a shared table deletion left nested cells visible.
- 3056f24: Table row borders can be dragged to set a row's height in both docx editors. The height saves as an at-least `w:trHeight`, and a drag is one undo step. Adds `set_row_height` to the engine.
- c74b2b8: The docx editing engine is smaller to download.
- c2e450c: Text boxes take the document's default East Asian language, so Japanese and Korean text boxes draw with their own CJK face. A font used by several CJK scripts puts the document's own script first.
- c87ef90: Text boxes apply all package paragraph and character styles by default, including inherited fonts, colors, languages and spacing. Shared state preserves these styles across peers and payload edits without changing saved payloads. `font_preflight` includes styled text-box faces and inherited language scripts.
- 0bee43d: Text runs drawn without shaping carry the script faces their paragraph needs (`scripts`), so the native renderer requests the Japanese, Korean or Traditional Chinese face the browser path already chooses.
- ad45f86: Adds opt-in `decodeFontFace` for packed fonts. Format package roots return a promise of `font` and optional `payload` ArrayBuffers; native bindings return byte arrays synchronously. Invalid options and decode errors reject or throw without retrying another decoder.
- b8395eb: `fontPackDiagnostics()` names installed font add-ons that cannot be used (another release, missing files, unreadable manifest). Native packages also warn once per pack with code `BETTEROFFICE_FONT_PACK`. Exports keep substituting those faces.
- 13a0dc0: EMF and WMF pictures that draw the same bitmap several times, or whose EMF+ drawing falls back to the GDI one, no longer exceed the bitmap budget: equal bitmaps are decoded and counted once, and a set-aside pass's bitmaps are released first.
- 92ca93c: EMF and WMF pictures that draw the same bitmap several times now rasterize within the image budget in native PNG exports: a repeated bitmap's decoded pixels are counted once, as the renderer decodes and frees each draw on its own.
- 79ae853: Native rendering now draws JPEG and PNG pictures past the pixel cap at a reduced size, matching the canvas downscale, instead of leaving them out. Pictures whose reduced decode would still exceed the memory budget are refused as before.
- 1af9969: Adds prebuilt binaries for musl Linux on x64 and arm64, such as Alpine, and for Windows on arm64.
- ec04baa: The native packages require Node.js 20.16+ or 22.3+, which they need to read installed font packs from disk.
- 681b02e: A Rust panic in a native engine now fails its session as a trap, as on WebAssembly, instead of aborting the process. The engine object that panicked refuses every later call except free.
- a121ec6: Native engine classes now have `[Symbol.dispose]`, the same function as `free()`, as on the wasm packages, so `using` declarations release them at scope exit.
- 6378188: Worker sessions now run on Node.js >= 20.16 / 22.3 through `worker_threads`, keeping the process alive until disposed. DOCX worker editing needs `OffscreenCanvas` and fails with a clear error without it (Node.js, Bun, Deno); open, read and save still work.
- f1e7906: Adds experimental `fontPreflight` and `preloadPreflightWasm` for the active slide and the opening sheet, without opening the editing engine. DOCX `fontPreflight` answers now follow the engine for notes, the main part, numbering, theme script fonts, style defaults and table columns.
- 76bb148: Native slide and page PNGs draw a PNG picture that ends without its closing chunk, as browsers do, instead of skipping it.
- 7cf70f8: Worker entries no longer import the module that constructs their own worker, so Next.js (Turbopack) production builds of apps that use `./worker` finish instead of hanging.
- 55c25ba: Worker files and main-thread entries import no private `#` specifiers, so worker sessions open from a CDN or with an import map, without a bundler. On Node.js, a worker thread that fails to start now fails the session instead of throwing synchronously.
- fd25d97: Text in the Symbol font now draws the Greek letters, arrows and math signs Office shows for it, in documents, slides and cells alike. Wingdings and Webdings run text uses the same shared slot table as bullets. Saved text is unchanged.
- 1175683: TIFF pictures whose decoded size their dimensions don't bound, such as JPEG-compressed ones, are no longer decoded, keeping memory use predictable. Adds `tiff_bounded_pixels`. DOCX opens charge a per-document TIFF pixel budget. VSDX rendering skips TIFF pictures.
- f2df1e8: Collaboration updates with invalid UTF-8 text or oversized length prefixes are refused with an error instead of being applied. The macOS app now builds the same yrs as every other target.
- 6d79fc5: Undo tracking no longer slows with the number of stories, so large deletions and clears in documents with many table cells stay fast. Published Rust crates get this once yrs ships it.
- Updated dependencies [f99e53a]
- Updated dependencies [de8f092]
- Updated dependencies [1a82d79]
- Updated dependencies [3609ce4]
- Updated dependencies [5999962]
- Updated dependencies [57f4975]
- Updated dependencies [d60b8e4]
- Updated dependencies [9ab2b61]
- Updated dependencies [39ae19a]
- Updated dependencies [b8395eb]
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
- Updated dependencies [9ac1ffe]
- Updated dependencies [ad45f86]
- Updated dependencies [e6d5c52]
  - @betteroffice/fonts@0.5.0
  - @betteroffice/fonts-cjk@0.5.0
  - @betteroffice/fonts-intl@0.5.0
  - @betteroffice/fonts-latin@0.5.0

## 0.5.1

## 0.5.0

### Patch Changes

- c688743: Adds prebuilt Rust engines for Node.js, Bun and Deno that open, edit, render PNGs and save documents. Image decoding stays within the render budget.
