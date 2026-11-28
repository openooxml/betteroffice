# @betteroffice/docx-native

## 0.6.2

### Patch Changes

- bee9d4d: Typing in a table cell of a document with one section no longer lays out the whole document again on most keystrokes. Documents with section breaks keep the full layout for cell edits.
- c9a261c: Editing next to text-wrapped floating objects in DOCX lays out faster, because layout after the edit stops once the following text is unchanged. Section grid changes beside such objects are respected.
- 94e792d: Reduces repeated work when updating selections, page objects, headers and footers while typing in long DOCX documents.
- ea9805e: Displays cached values from nested DOCX fields in document text.
- 5f07dd7: Keeps tables of contents and other fields that span paragraphs intact when their entries are edited, inserted or deleted before saving a DOCX file, so page numbers stay updatable in Word.
- 504f163: Word and PowerPoint draw corrected bent connector paths and preserve dash alignment and opacity in PNG exports. Word arrowheads follow the connector’s direction, and PowerPoint bent connectors are selectable along the path, including mitered corners.
- 305d381: A font picked for Chinese, Japanese or Korean text in a document now stays after saving and reopening in Word, and Latin text keeps its East Asian font. Spreadsheet cells given a font no longer follow the theme font in Excel.
- 8475b31: The Linux GNU binaries load again on glibc 2.17 and newer on both supported architectures.
- a42798b: Documents saved on Windows are now byte-identical to the same documents saved on Linux and macOS. Saved content does not change.
- Updated dependencies [e163044]
- Updated dependencies [662a971]
  - @betteroffice/fonts@0.5.1
  - @betteroffice/fonts-cjk@0.5.1
  - @betteroffice/fonts-intl@0.5.1
  - @betteroffice/fonts-latin@0.5.1

## 0.6.1

### Patch Changes

- 365ae9d: The Linux native packages load again on Ubuntu 24.04, Debian 12 and other distributions with glibc 2.35 or newer.

## 0.6.0

### Minor Changes

- 7db6627: Created-list saves no longer fail or write wrong packages for differently spelled relationships or numbering parts, keep ids clear of the document's, and write only definitions in use. The list ensure keeps levels and joins no neighbour's style list. New: `list_definition_keys_in_use`, `list_definition_keys_in_use_json`, `listDefinitionKeysInUse`.
- ccb79e9: Text inside `w:dir` and `w:bdo` is shown, found and kept on save and edit. TypeScript: runs and `ComplexField` carry optional `directionScopes`. Rust: `Run` and `ComplexField` gain public `direction_scopes` (struct literals and exhaustive patterns must add it); `DirectionScope` and `DirectionKind` are exported.
- 78f28c5: `applyEdits` text steps and new paragraphs accept `runs` with bold, italic, underline, strike, colour, highlight, size, font and superscript/subscript marks instead of plain `text`. Rust `EditOperation` and `ParagraphInput` gain a `runs` field.
- 2a1d763: Adds the `formatText` edit step (`set`, `clear` back to the styles, `onUnsupported`). The agents `format_text` tool gains strike, superscript/subscript, colour, highlight, size and font. `EditReceipt` literals must set `skipped`.
- ccb79e9: Unmodelled `w:` inline elements are kept on save. Parse warnings gain `DOCX_UNMODELED_INLINE` for kept ones holding text, such as tracked changes inside a hyperlink, which an edit of their paragraph still drops. Rust: `RawOp` gains `RawOp::SetRawMark`, so exhaustive matches must handle it.
- 09b46fc: Bullets and Numbering create a list of their kind or join the one directly above or below, put a selection on one list, and lay it out as saved. The engine gains `ensure_list_definition_after` and `paragraph_style_numbers`; `ensure_list_definition` also returns `listLevelIndents`; an engine without `ensureListDefinition` refuses `toggleList`.
- 57fa9d7: Screen reader page mirrors read a visible list bullet as a localised word, never its stored character. `DocAttrs` gains `listMarkerBullet` and `MirrorLabels` gains `bullet`; without either, a bullet stays hidden. Rust `DocAttrs` literals must set `list_marker_bullet`.
- b1a4b01: Screen reader page mirrors read list numbers and labels such as "3.2" or "Appendix 1", which were hidden. Bullets stay hidden. Rust `DocAttrs` literals must set `list_marker_numbered`.
- c9a8850: The native `EditSession` now also lays out, builds display frames, measures fonts, takes resident input, hit tests and reports update callbacks, and every WebAssembly engine function is exported with the same name and results.
- 590615b: The native DOCX engine can report its editing heap per thread. Counting is off by default; `startHeapCounting()` from `@betteroffice/docx/yrs` turns it on. Pressure events on native need a host memory budget.
- 905ef77: The raw `EditSession.renderPngTarget(displayListJson, targetJson, optionsJson)` rasterizes one laid-out page to PNG in the native engine, from the same display list the canvas paints, with the session's fonts. Oversized pages and missing fonts are refused with stable messages.
- e9fbbe6: `@betteroffice/docx-native` replaces its facade with `@betteroffice/docx`'s API and entry points, using native execution by default. Node.js, Bun and Deno imports need no WebAssembly initialization or assets. Replace `openDocument`, `Document` and `DocxDocument` with `parseDocx` and drop `DocumentStructure`, `EditReceipt`, `FontFace`, `ImageResource`, `LayoutResult`, `OpenDocumentOptions`, `RenderedPage` and `SaveDocumentOptions` imports.
- 129e1eb: A complex field held in one run keeps its cached result and instruction spacing when its paragraph is rewritten, also after seeding resolves its theme fonts. New: the parser's public `docx_parse::field_run_digest` and the engine's `field_run_digest`.
- c0e297a: The engine reads the text a laid-out page shows, built or not: `page_texts_json` returns each page's paragraphs, cut where the page cuts them, with heading levels, and the table rows the page places.
- a9897ef: DOCX PNG export tells onDiagnostic, once per page, which kinds of effect the engine does not draw (`effects-skipped`). The shared PNG diagnostic type gains that code, so a host with an exhaustive switch on the code gets a new case.
- 62605af: DOCX PNG export draws a placeholder for a picture it cannot draw instead of rejecting the page or leaving it blank; `onDiagnostic` reports each one and why. `strictPictures: true` keeps rejecting. Agent previews warn.
- e20cdb3: `YrsSession.renderPng({ page }, options?)` draws a page as PNG bytes with the editor's layout and painter, or natively with `@betteroffice/docx-native`. It rejects with `PngExportError` rather than paint a substitute font.
- 657aaea: A field spanning paragraphs no longer repeats the later paragraphs inside itself: its structured code and result carry `blockCount` (0: not tracked). Read those paragraphs from the story, not `blocks`. Rust code building `StructuredFieldContent` must add the new `block_count` member.
- ae6e7e1: Fixes collaborative documents that stopped laying out after an undo removed a table cell another collaborator's table still showed. That cell now reads as empty and refuses edits until the undo is redone or its row is deleted. Handle `EditError::CellHasNoStory` and `OpError::CellHasNoStory` in exhaustive matches.
- 6d8fb19: A table that crossing undos left unlayoutable is now skipped instead of breaking layout; table commands and saves refuse with a coded error until redo restores it or it is deleted. Both engines add `unavailable_tables_from`. Exhaustive matches must handle `OpError::TableUnavailable`.
- b2ea7f5: The docx and pptx native bindings use mimalloc as their memory allocator, through the default-on `mimalloc` cargo feature. DOCX save, edit and layout get faster and PPTX open and export faster, with a higher peak memory on some documents.
- cf72bb4: The native packages now export the full `@betteroffice/docx`, `@betteroffice/xlsx` and `@betteroffice/pptx` APIs as ES modules, and `require()` of them needs Node 20.19 or 22.12: `engines` now says so. `@betteroffice/agents`, which depends on them, requires Node 22.12.
- f78484e: Adds `prime_last_slice_cost` to the wasm and native DOCX edit sessions: what the last `prime_save_seed_states_slice` call charged against its byte budget, overshoot included, or 0 (committed-scan slices charge nothing). The docx worker sizes the next slice from it, growing at most twofold.

### Patch Changes

- 1282f77: Adds opt-in `set_language` edits for DOCX paragraphs, grep matches, styles and document defaults. Latin, East Asian and complex-script proofing language slots can be set or cleared independently while preserving other formatting. Rust `InlineFormatDelta` literals must set `language`.
- 9590d1e: Paragraph-relative shapes now move with an anchor paragraph whose first line overflows to the next page, instead of staying behind on the previous page.
- 232fe11: Text set in Simplified Arabic or Traditional Arabic now uses the line height Word uses for those fonts, so Arabic documents break pages closer to where Word breaks them.
- 27130ef: An autofit table whose explicit width is wider than the text area now keeps that width instead of being narrowed to the text area, matching Word. Percentage and automatic widths are unchanged.
- c3d6f21: Autofit tables with no table width keep their cells' stated widths when those widths and the column grid both exceed the text width, as Word does, instead of squeezing the columns and wrapping text that fits.
- c0ae1af: Formatting a paragraph mark inside a block content control, such as a form field, no longer rewrites the rest of the document on save: tracked changes, proofing marks, drawings and fields elsewhere are kept as they were.
- f9c8379: Typing in a body paragraph that directly follows a table or column break is faster: each keystroke now updates only that paragraph instead of the whole body.
- e1a414e: Documents whose XML parts start with a byte order mark and have bookmarks between paragraphs now save. Before, saving them failed with a malformed XML error.
- c2b9ac7: Tables inserted or pasted in the editor lay out with the cell margins of the document's table styles, as Word shows them once saved, so text wraps on screen where Word wraps it. Opened tables are unchanged.
- 675b859: Most typing in a table cell of a long document now lays out only that table instead of the whole body, and tables after the edit keep their measurements instead of being remeasured.
- 455254f: Native DOCX rendering no longer fails on pages with charts or comma-separated font names. Chart text naming a theme font now uses the document theme's font in every engine.
- fedbb82: Clear Formatting in the DOCX editor returns text to what its paragraph and table styles give, so headings stay bold. `formatText` clear `'all'` no longer changes a run's text direction. Paragraph patches gain an optional `inheritance` field. Rust `ParagraphPatchV1` literals must set `inheritance`.
- 3c1c056: Native DOCX page rendering is up to ten times faster on pages with clipped content. Pixel output is unchanged.
- 8ee77cd: A comment whose range sits inside a tracked change, a hyperlink or a content control no longer loses that range when another paragraph is edited and the save keeps the source paragraph. Deleting such a comment removes its markers.
- 54a0219: The DOCX engine holds list definitions the editor creates and writes them on save. Documents without one are unchanged. New in the edit crate: `ListKind`, `ListIdentity`, `CreatedListDefinition`, `resolve_list_identity`, `ensure_list_definition` and `created_list_definitions`.
- cd6e71d: Deleting a comment in the editor and saving no longer forces a rewrite of every part. Where the writer can patch the body, a comment inside one paragraph leaves the other paragraphs' source bytes unchanged. Comment and bookmark ids differing only by leading zeros match.
- 46c17cb: Documents whose styles declare empty document defaults lay out their unstyled paragraphs and table cells without paragraph spacing, as Word does, instead of with Word's 8 pt after and 1.08 line spacing.
- c13d981: A document whose last paragraph holds only a page break now starts its endnotes under the last line of text, as Word does, instead of on a new page.
- 33fcb87: Endnotes now follow the last body text, as in Word, and continue onto new pages under a continuation separator. Before, long endnotes pushed body text onto extra pages and were painted over it. Set new fields in `NoteAreaContract`, `NoteRegion`, `ReservedBlock` and `NoteLayoutItemContract` literals.
- f6946f9: Pictures in equal-width columns of a continuous section that keeps its side margins are placed as in earlier releases again, and an edit in the section before no longer moves them.
- a9b74ab: Preserves bookmarks inside complex field results when saving, including fields that span paragraphs in the document body or table cells.
- 363ceb2: One run holding two or more whole fields keeps its fields and the text between them on save while the field is unedited: `ComplexField` gains optional `sourceRunXml` and `sourceRunDigest` (`source_run_xml`, `source_run_digest` in the Rust struct; struct literals must add both).
- 0a4285c: Adds opt-in `fontPreflight`, returning DOCX font faces, script keys, and scan completeness without the editing engine. Its standalone module loads only when called or explicitly loaded with `preloadPreflightWasm`.
- c3c5294: Picture watermarks that Word also places in a footer are kept when the document is saved.
- c2281b5: Formatting selections spanning tables (select-all then bold) formats cell text. Paragraphs that would lose content on save stay unchanged. Edits and saves report which and why. Have `YrsSession` hosts and mocks return `YrsFormatReceipt` from `toggleMark`, `formatRange` and `clearFormatting`, and initialize new `S13ParagraphIds` and `S13SplicedPart` fields.
- fa2717d: Every engine now reads decimal numbers back exactly. Drawing groups whose scaled child coordinates have long decimals stay intact on save.
- 6a5dded: Word text watermarks now draw on the pages whose header carries them and no longer on a page whose header has none. Rust literals must set `HeaderFooter.watermark_box`, `HeaderFooterVariant.watermark` and `header_watermarks` on `HeaderFooterPayload` and `DocumentRegions`.
- 8e4260c: A DOCX table row marked hidden keeps its height when any paragraph mark in it is visible, as Word does, even if the paragraphs are empty. Only rows whose paragraph marks and runs are all hidden are left out.
- fc50a58: Unless hidden text is shown, DOCX rows marked `w:trPr/w:hidden` with entirely hidden content occupy no height and paint neither borders nor shading. Marked rows with visible content render. Fully hidden tables take no space. Cells retain document row numbers. Initialize new `TableRow` and `TableRowFormatting` fields.
- 869c704: An inline SmartArt stays in its paragraph's flow, between the text around it, instead of being laid out above the paragraph. Saving an edit after a page or column break that follows the diagram keeps the break after it.
- b756b5f: Runs whose character scale is 0, which Word draws at full width, now draw at full width instead of disappearing on canvas and failing the page natively. Other values outside 1 to 600 are ignored or held to 600.
- 773c633: Keeps bookmarks in table cells, content controls, headers, footers, notes, comments and text boxes on save, including after blocks are deleted, merged or moved. Obtain `XmlElement` from `parse_xml` and read retained bookmark markup from `bookmark_replay_xml`.
- 8efeef9: Keeps bookmarks inside field codes, such as the one Word writes for each legacy form field, and fields nested inside another field's code, on save.
- 51a1bbf: Keeps drawing groups and canvases on save in headers, footers and rewritten paragraphs. A group that matches its source apart from its frame and text-box text is written back from the source with those changes; any other edit writes it from the model as before.
- 30134ef: Keeps bookmarks that sit between table rows or between cells on save, including after rows or columns are inserted or deleted in an editing session. Rust literals must set `trailing_bookmarks` on `Table`, `leading_bookmarks` on `TableCell`, and both on `TableRow`.
- bf8880c: Keeps bookmarks inside the paragraphs of text boxes on save, including the copies Word stores for older versions.
- f777004: Keeps a text box on its own run on save when bookmarks, comment ranges or other markers come before it in the paragraph; such a text box was dropped or moved to another run.
- f31b62e: Keeps text boxes inside tracked insertions, deletions, hyperlinks and simple fields on save, with their text and revision marks.
- 002ac93: An open session now repaginates like a freshly opened one after edits next to a page break or to a continuous section's margins. `PageFlowGeometry` can no longer be built with a struct literal outside docx-layout.
- 3789492: A list marker larger than its hanging indent now starts at the hanging indent, as in Word, instead of further left.
- afcb67e: A DOCX list level override that sets no start value now keeps the start of the level it overrides, so its first label reads "A." or "3." instead of nothing or "0.".
- 10fbc38: DOCX numbered list levels without a start value now count from zero, as Word does, and a zero counter prints "0." instead of nothing.
- fa27c3c: DOCX lines measured with authoritative shaping now carry the glyphs measurement shaped them with, as `glyphRuns`, packed into one string per line. Rust `TypesetRow` literals must set `glyph_runs`.
- 464d5c6: Screen reader page mirrors read a chart as a figure holding its values as a table, series as column headers and categories as row headers, formatted like its data labels. A chart without cached values says it has no data. `ChartA11yAttrs` literals must set `data`.
- 59d467a: The accessible page text reads a line break inside a paragraph as a new line and a tab as a tab, so the words on either side no longer run together for screen readers. Rust `DocAttrs` literals must set `separator_before` and `separator_after`.
- 140c8fa: Screen reader page mirrors read the first column of a table with "Repeat as header row" rows and first-column formatting as row headers, and data cells name their row header. Column headers are unchanged. Rust literals must set `TableCellRef.is_row_header`, `TableCell.merge_continuation` and `TableBlock.first_column`.
- 29dc450: Text with Word's classic shadow draws it instead of an empty Word 2010 shadow, a Word 2010 shadow no longer adds a classic one and keeps its own settings, and a content control's data binding is no longer replaced by its Word 2013 copy.
- 57f4975: Arabic text set in Simplified Arabic or Traditional Arabic now falls back to the bundled Naskh face, closer to those fonts' style and width.
- 452e0a9: Native DOCX PNG export paints rectangular and path gradient fills as a radial gradient, as the browser renderer does, instead of failing the page.
- 533fb57: The native DOCX engine now enforces the worker heap limit: a call past it fails with the same out-of-memory error as wasm, and its object refuses later calls. A limit turns heap counting on; without a limit or `startHeapCounting()`, counting costs nothing.
- dbcc3fd: Undo grouping reads JavaScript `Date.now`, as the wasm engine does, so edits group the same way on both engines, including under a mocked clock.
- 2dbba20: `@betteroffice/docx-native` installs the Latin, international and CJK font packs and reads each default face from disk on first use, so international and CJK text gets its bundled face with nothing fetched.
- 4520944: `install_panic_hook` now works and the package installs it on load: on macOS and Linux a Rust panic writes one `[docx-native] panic:` line to stderr, then Rust's default report. The thrown error is unchanged; calling it again does nothing.
- 56c2c6c: Native docx page PNGs draw SVG, EMF and WMF pictures. A picture the renderer still cannot draw rejects `renderPng` with the new `picture-unavailable` code instead of leaving it blank.
- ce686a8: Footnote and endnote reference marks in the document view now print the note's number, in its format and from its starting value, as the note area shows it, instead of its internal id, also while the per-revision accept/reject preview hides or shows a reference.
- fa08d5d: Ligatures such as fi require document opt-in, matching Word. Contextual alternates follow the document too. The caret and selection land between a ligature's letters, and selecting a field result covers all of it. Set new fields in `TextFormatting`, `RunFormatting`, `DocAttrs` and `DisplayTextContractMetadata` literals.
- 3a0020d: Typing in an empty paragraph that is followed by a column break and a page break no longer adds an extra page.
- ed0a3ce: DOCX Arabic, Syriac and Devanagari now paint the glyphs measurement shaped, so words join on screen, including words split across runs that differ only outside the complex-script font.
- 934104f: A table pasted from another document lays out with this document's Word compatibility mode, as Word shows it once saved, instead of the source document's, so it no longer shifts left by its cell margin.
- 22fc562: Clearing or changing character spacing, scaling, kerning or raised text, and clearing a theme font, no longer rebuilds the paragraph on save, so its hyperlinks, smart tags and drawings are kept as they were.
- 7796969: Typing next to a field whose begin, code, result and end all sit in one paragraph no longer rebuilds that paragraph on save, so its other content is kept as it was.
- 101fbe9: Documents with a picture beside text in equal-width columns of a continuous section lay out again instead of failing.
- 914dabc: DOCX PNG export names a picture whose part is missing by its relationship id, in `onDiagnostic` and in the `strictPictures` rejection, instead of leaving it unnamed. Display-list pictures carry it as `relationshipId`. Rust `ImageRun`, `ImageBlock` and `ImagePrimitive` literals must set `relationship_id`.
- 6d54917: Native page PNGs clip elliptical pictures and their borders to the ellipse, as the editor does, instead of drawing a rectangle.
- a2d41a8: Native page PNGs draw picture-filled shapes, stretched or tiled, with their outlines, as the editor does, instead of leaving the shape out.
- aa54bb8: Native page PNGs accept VML color names, and a list marker in the paragraph's default font no longer refuses the page with a missing font chain.
- 857b6ea: README corrections: headless editing is `createYrsSession` from the package's `yrs` entry and `session.openDocx`, not the unexported `EditSession`. In the docx README, `@betteroffice/fonts` comes with the package and only the CJK and international packs are optional.
- 94093ba: `parse_docx_s4` no longer fails on an unreadable relationships part other than the document's own. It reports a warning and leaves that part's references unresolved. Rust `S4Projection` literals must set `warnings`.
- 524652d: A page or column break that an edit moves to the start of the next paragraph is now saved there instead of being dropped, in paragraphs that start with a plain run.
- 55ef1d4: Saving an edited paragraph keeps a page break before, a text direction and a snap to grid setting that are explicitly switched off.
- 1a76823: Saving keeps the paragraphs beside a comment that ends at the start of a paragraph unchanged.
- cfdeb9e: Saving a document after an edit that changes many paragraphs, such as clearing formatting over the whole body, is much faster on large documents. Saved files are unchanged.
- 365b54f: The DOCX paragraph save plan now lists `touchedParts` sorted by part path, with each part's paragraph ordinals ascending, so the same edits give the same plan in every session. Saved files are unchanged.
- cda31aa: Saving a document records paragraph ids in linear time instead of comparing every saved id with every live paragraph. At 5 MiB this step drops from about 2.7 s to 63 ms on native and 4.9 s to 0.2 s on wasm; output is unchanged.
- a711586: When a DOCX is first loaded for editing, equal attribute keys and values share one allocation. Memory retained after open drops by 1 to 6% on large real documents (10% on a text-heavy synthetic 5 MiB file). Peak memory and saved bytes are unchanged.
- 84b271b: Opening a DOCX keeps equal embed keys and values in one shared allocation: retained memory after open falls 7.8 to 12.0% on the two largest public test documents and 29.2 to 30.5% on a 5 MiB synthetic document, native and wasm. Saved bytes are unchanged.
- 6f65928: Fixes comment and tracked-change ID reuse when reopening with the same `clientId` from a checkpoint containing the client's earlier comments and tracked changes.
- 3789492: Shapes, SmartArt and shape text now use the document theme's colours instead of the Office 2013 palette. Shape outlines and SmartArt text without their own colour use their style's line and font colours.
- 946f6f5: Picture fills in SmartArt now resolve against the diagram drawing's own relationships, so they draw instead of failing the page natively and falling back to a solid fill on canvas. Pass a media map or `None` before the context in `parse_smart_art_from_drawing`.
- 3789492: Each shape of a SmartArt diagram shows its own text, and shapes without text stay empty, instead of every shape repeating the last shape's text.
- 3789492: SmartArt text now sits in its own text box with Word's rotation, bullets, spacing and colour-scheme text colours. Shape fills, lines and shadows keep their transparency and size, line arrowheads draw, and more preset shapes render.
- 35a4fc2: Editing the first or last paragraph of a field that spans paragraphs, such as a table of contents, no longer saves the field without its end or with one end too many.
- 514e876: Story checksums use constant memory, so long documents no longer exhaust the WebAssembly engine's memory and stop the editor. Checksum values are unchanged.
- bbac3e7: Fixes collaborative table editing where splitting a cell again after a collaborator removed the previous split cell made the new cell invisible, and where redoing a shared table deletion left nested cells visible.
- 3056f24: Table row borders can be dragged to set a row's height in both docx editors. The height saves as an at-least `w:trHeight`, and a drag is one undo step. Adds `set_row_height` to the engine.
- c74b2b8: The docx editing engine is smaller to download.
- c2e450c: Text boxes take the document's default East Asian language, so Japanese and Korean text boxes draw with their own CJK face. A font used by several CJK scripts puts the document's own script first.
- c87ef90: Text boxes apply all package paragraph and character styles by default, including inherited fonts, colors, languages and spacing. Shared state preserves these styles across peers and payload edits without changing saved payloads. `font_preflight` includes styled text-box faces and inherited language scripts.
- 0bee43d: Text runs drawn without shaping carry the script faces their paragraph needs (`scripts`), so the native renderer requests the Japanese, Korean or Traditional Chinese face the browser path already chooses. Rust `TextRunPrimitive` literals must set `scripts`.
- 6f381d7: A line break or a comment anchor inside a tracked change was lost from the saved file when its paragraph was edited.
- d55af86: A footnote or endnote reference lost its revision when its paragraph was edited: a deleted reference came back as live. It now keeps its revision, formatting beside it is usually saved, and one suggested for deletion inside a field result stays in the file.
- 0f9f39e: In DOCX, the mark after a tracked inserted or deleted paragraph is drawn from the document's own default-family face when the document uses it and that face is loaded whole, and by the browser otherwise.
- 675b859: After each keystroke in a long document, tracked changes are read again only for the parts of the document the edit changed, and content controls and page headings are rebuilt only when something they depend on changed.
- 16316e4: Typing in a document body no longer briefly shows a line with its characters shifted.
- c0fc51b: Text of grouped shapes inside table cells with vertical text is visible again instead of being clipped away.
- fe023f2: Negative luminance offsets now darken DrawingML colours.
- ad45f86: Adds opt-in `decodeFontFace` for packed fonts. Format package roots return a promise of `font` and optional `payload` ArrayBuffers; native bindings return byte arrays synchronously. Invalid options and decode errors reject or throw without retrying another decoder.
- b8395eb: `fontPackDiagnostics()` names installed font add-ons that cannot be used (another release, missing files, unreadable manifest). Native packages also warn once per pack with code `BETTEROFFICE_FONT_PACK`. Exports keep substituting those faces.
- 13a0dc0: EMF and WMF pictures that draw the same bitmap several times, or whose EMF+ drawing falls back to the GDI one, no longer exceed the bitmap budget: equal bitmaps are decoded and counted once, and a set-aside pass's bitmaps are released first.
- 92ca93c: EMF and WMF pictures that draw the same bitmap several times now rasterize within the image budget in native PNG exports: a repeated bitmap's decoded pixels are counted once, as the renderer decodes and frees each draw on its own.
- eb5c453: DOCX and PPTX native sessions can release their observers and memory after an engine failure. The DOCX worker reports failures promptly and passes along queued collaboration updates.
- 79ae853: Native rendering now draws JPEG and PNG pictures past the pixel cap at a reduced size, matching the canvas downscale, instead of leaving them out. Pictures whose reduced decode would still exceed the memory budget are refused as before.
- 1af9969: Adds prebuilt binaries for musl Linux on x64 and arm64, such as Alpine, and for Windows on arm64.
- ec04baa: The native packages require Node.js 20.16+ or 22.3+, which they need to read installed font packs from disk.
- 7f439b2: The native packages and each of their platform packages now ship the project's `LICENSE` file in the tarball, as the WebAssembly packages do. They declared Apache-2.0 without carrying the text.
- 681b02e: A Rust panic in a native engine now fails its session as a trap, as on WebAssembly, instead of aborting the process. The engine object that panicked refuses every later call except free.
- 8a17027: PNG export on the native engine can reduce memory retained by the exporting session.
- 4687b6f: `require()` of the native packages from CommonJS resolves, as it does for the WebAssembly packages: every export gains a `default` condition beside `import`. Before, it failed with `ERR_PACKAGE_PATH_NOT_EXPORTED`. Loading an ES module by `require()` needs Node 20.19+ or 22.12+.
- a121ec6: Native engine classes now have `[Symbol.dispose]`, the same function as `free()`, as on the wasm packages, so `using` declarations release them at scope exit.
- 9e9309a: Each worker loads as one file.
- 6378188: Worker sessions now run on Node.js >= 20.16 / 22.3 through `worker_threads`, keeping the process alive until disposed. DOCX worker editing needs `OffscreenCanvas` and fails with a clear error without it (Node.js, Bun, Deno); open, read and save still work.
- d56440d: In Safari and WebKit, apps that bundle the editors with ES module workers (Vite `worker.format: 'es'`) could run a worker twice and fail on its first command. Each worker entry now loads its code once.
- f1e7906: Adds experimental `fontPreflight` and `preloadPreflightWasm` for the active slide and the opening sheet, without opening the editing engine. DOCX `fontPreflight` answers now follow the engine for notes, the main part, numbering, theme script fonts, style defaults and table columns.
- 76bb148: Native slide and page PNGs draw a PNG picture that ends without its closing chunk, as browsers do, instead of skipping it.
- d1166bc: `Presentation::save` in Rust and `Presentation.save` in Python succeed after slides are deleted and prepare fonts for added slides. DOCX layout limits a section's column count to what Word supports, so a very large count no longer stalls `Document::layout`.
- 7cf70f8: Worker entries no longer import the module that constructs their own worker, so Next.js (Turbopack) production builds of apps that use `./worker` finish instead of hanging.
- 55c25ba: Worker files and main-thread entries import no private `#` specifiers, so worker sessions open from a CDN or with an import map, without a bundler. On Node.js, a worker thread that fails to start now fails the session instead of throwing synchronously.
- fd25d97: Text in the Symbol font draws the Greek letters, arrows and math signs Office shows for it, in documents, slides and cells alike. Wingdings and Webdings run text uses the same shared slot table as bullets. Saved text is unchanged. Handle `SymbolFont::Symbol` in exhaustive matches.
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
- Updated dependencies [f52a44a]
- Updated dependencies [b8395eb]
- Updated dependencies [d60b8e4]
- Updated dependencies [53502c6]
- Updated dependencies [58a8ca6]
- Updated dependencies [ec04baa]
- Updated dependencies [d60b8e4]
- Updated dependencies [21a60e5]
- Updated dependencies [6aefa34]
- Updated dependencies [19c20c3]
- Updated dependencies [df33f6a]
- Updated dependencies [df33f6a]
- Updated dependencies [37199d9]
- Updated dependencies [842d624]
- Updated dependencies [e65df1c]
- Updated dependencies [18b3628]
- Updated dependencies [dc78510]
- Updated dependencies [54d400a]
- Updated dependencies [ad45f86]
- Updated dependencies [e6d5c52]
- Updated dependencies [9e9309a]
  - @betteroffice/fonts@0.5.0
  - @betteroffice/fonts-cjk@0.5.0
  - @betteroffice/fonts-intl@0.5.0
  - @betteroffice/fonts-latin@0.5.0

## 0.5.1

## 0.5.0

### Patch Changes

- c688743: Adds prebuilt Rust engines for Node.js, Bun and Deno that open, edit, render PNGs and save documents. Image decoding stays within the render budget.
