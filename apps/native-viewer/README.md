# BetterOffice native viewer

This experimental macOS-first app paints BetterOffice DOCX, XLSX, and PPTX display lists with Vello.

From this directory, render and compare the first DOCX page:

```sh
cargo run --release -- --png page.png
```

Render the used range of the first sheet in the XLSX showcase:

```sh
cargo run --release -- --document ../demo/public/showcase.xlsx --sheet 1 --png sheet.png
```

Render the second slide in the PPTX demo:

```sh
cargo run --release -- --document ../demo/public/betteroffice-demo.pptx --slide 2 --png slide.png
```

DOCX and XLSX write the requested PNG through Vello and a sibling `.raster.png` through the matching existing raster backend, then print image-difference metrics. PPTX has no raster backend, so it writes only the Vello PNG and prints a JSON summary with primitive translation counts, skip reasons, positioned-glyph counts, and a caret-stop audit.

Use `--document FILE` and `--scale N` for any format. `--page N` selects a one-based DOCX page, `--sheet N` selects a one-based XLSX sheet, and `--slide N` selects a one-based PPTX slide. XLSX defaults to the selected sheet's used range and resolves supported charts from the workbook package. Every display-list item that cannot be translated is replaced by a magenta box and cross, counted by type, and reported with its reason.

Open the interactive viewer:

```sh
cargo run --release
```

To collaborate on one DOCX, XLSX, or PPTX, open the same file and room ID in two terminals:

```sh
cargo run --release -- --document ../demo/public/betteroffice-demo.docx --room native-demo
cargo run --release -- --document ../demo/public/showcase.xlsx --room native-sheet
cargo run --release -- --document ../demo/public/betteroffice-demo.pptx --room native-slides
```

Run that command once in each terminal. The default hosted relay can be replaced with `--relay-origin http://127.0.0.1:8787` or the `BETTEROFFICE_RELAY_ORIGIN` environment variable. The status line reports connecting, synced, peer count, and reconnect failures. Local edits made while disconnected remain in the document and are sent when the connection returns.

Collaboration synchronizes the resident DOCX, XLSX, or PPTX state. Native body-text, supported character or paragraph formatting, cell edits, and slide-text edits are broadcast; remote updates relayout the affected document, repaint the selected sheet, or repaint the changed slide. File paths, saved package bytes, viewport, local selection, and undo history are not shared. Awareness and presence are not wired: the native viewer answers awareness queries with an empty update and does not show or publish participant names, carets, or selections.

Run the native-to-TypeScript collaboration test from the repository root:

```sh
bun test apps/native-viewer/tests/native_typescript_collaboration.test.ts
```

The tests build the DOCX, XLSX, and PPTX Wasm artifacts when needed, start `apps/relay` locally with Wrangler, connect a headless native viewer peer through `--room`, and drive each format's production `CollaborationProvider`, engine replica, protocol implementation, and demo room transport directly from Bun. They prove retained-state joins, interleaved native and TypeScript edits, bidirectional offline edits after reconnect, identical canonical checksums, and identical state vectors. They are not part of the automatic root test run because they start a live relay and compile the native viewer.

This exercises the exact TypeScript collaboration module imported by the browser without UI behavior obscuring it. It is not browser-verified: it does not exercise the React editor's provider hookup, worker lifecycle, browser WebSocket implementation, or collaboration UI.

The fixed top toolbar provides undo, redo, bold, italic, underline, paragraph alignment, zoom, and save controls. A fixed status line shows the file and current page, sheet, or slide, plus transient save results. XLSX and PPTX enable save and their engine-backed undo and redo while keeping the DOCX character and paragraph controls disabled. Scroll vertically with the trackpad or mouse wheel. Hold Command or Control while scrolling to zoom. Use Command or Control with `+`, `-`, and `0` to zoom while editing.

DOCX documents are editable in the window. Click or drag to place a caret or select text, Shift-click to extend a selection, and double-click to select a word. Arrow keys, Home, End, typing, Backspace, Delete, and Enter use the resident DOCX engine. Command-Z or Control-Z undoes and adding Shift redoes. Press Command-S or Control-S to save beside the input as `<name>-edited.docx`; the viewer verifies the saved file by reopening it without replacing the live editing session. Save results and fidelity refusals appear in the status line.

XLSX worksheets are also editable. Click a cell or use the arrow keys to move the green selection border; the status line shows its address and current value. Typing replaces the cell, while Enter or F2 opens its current input. While editing, Enter commits and moves down, Tab commits and moves right, Escape cancels, and Backspace removes the last draft character. Every commit uses `betteroffice_xlsx::Workbook::edit_cell`, including its dependency-graph recalculation, then rebuilds the current sheet display list. Command-Z and Shift-Command-Z use the workbook undo and redo history. Command-S writes `<name>-edited.xlsx` beside the source and verifies it without replacing the live workbook.

PPTX text boxes are editable without range selection. Click text to select the layout engine's exact caret stop; typing, Backspace, Delete, Enter, Home, End, and the arrow keys stay within that text box and relayout the slide after every edit. Escape or a click outside text leaves editing. Command-Z and Shift-Command-Z use the presentation undo and redo history. Command-S writes `<name>-edited.pptx` beside the source and verifies it without replacing the live presentation or overwriting the original.

DOCX saving round-trips text edits, bold, italic, underline, and alignment in supported top-level body paragraphs. The save patches only edited paragraph XML into the source package, so untouched paragraphs and unrelated package parts remain byte-for-byte unchanged. Original paragraph properties are spliced back around projected edits, preserving unmodelled children such as `snapToGrid`, `textAlignment`, `wordWrap`, and `kinsoku`; an alignment edit changes only the original alignment child. Paragraph splits and merges disable Save with an explanation until they are undone. Saving also refuses nested table or content-control paragraph edits, combined text-and-character-format changes it cannot project safely, and any edit to paragraphs containing hyperlinks, comment or note references, revision marks, drawings or embedded objects, content controls, bookmarks, fields, math, or other non-text run content. The refusal identifies the paragraph and the markup at risk before writing an output file.

Successful saves retain the current caret or cell selection, scroll position, and undo history. Escape and the window close control refuse to close while committed edits remain unsaved; an active XLSX draft still uses Escape as its explicit cancel action.

XLSX saving uses the engine's preserved-package serializer. Worksheet parts whose modeled cells and recalculated caches did not change, along with unrelated sheets and package parts, remain byte-for-byte unchanged. Styles, rich shared-string entries and their source identities, defined names, drawings, charts, tables, conditional formatting, validation, and pivot parts are retained; a changed workbook drops its stale calculation chain and requests a full calculation on load. Affected worksheet parts reserialize modeled columns, rows, cells, formulas, and merges while retaining the other worksheet children. Before writing, the viewer refuses an affected worksheet containing column, row, cell, formula, merge, or extension markup that projection cannot reproduce. The refusal specifically identifies shared, array, and data-table formulas, rich inline strings, metadata, or extension markup at risk. Digitally signed packages are refused because any edit would invalidate the package signature. The engine's own package preflight remains authoritative for chart and pivot references it cannot rewrite, and an app-level audit refuses any unexpected change to an unrelated part. No output file is touched when either check refuses the save.

PPTX saving uses the engine's source-part patcher. Untouched slides and every unrelated package part must remain byte-for-byte unchanged. For an edited slide, the viewer permits only the edited shape's paragraph XML to differ, reopens the result to compare the full deck model, and maps unchanged UTF-16 story tokens back to their source run and paragraph XML. Language, hyperlinks, fields, theme-linked colors, paragraph properties, and unmodeled XML attached to unchanged text must survive exactly. Saving refuses edits that cannot be isolated to a source shape, multiple edited stories in one shape, package signatures, unexpected part or shape changes, metadata flattening or loss, and stories too large for the exact metadata audit. The refusal names the shape and property at risk, and no output file is touched.

The DOCX translation covers positioned glyph runs, fallback text shaping, rectangles, line and shape paths, embedded data-URL and scoped relationship images with crop/flip/rotation, and decorations. Advanced DrawingML effects and paint, filtered or bordered images, secondary-color lines, and compound or wave borders remain explicit skips.

The XLSX translation covers fills, clipped solid/dashed/dotted/double lines, geometry paths, and Carlito-shaped text with alignment, synthetic bold/italic, highlight, underline, strike, and dashed underline. Font-family requests intentionally follow `xlsx-raster` and use its bundled Carlito fallback.

The PPTX translation uses the slide layout engine and replays its positioned glyphs with the registered Arial-compatible faces. It covers positioned glyph runs, underline, normalized shape paths, solid and linear/radial gradient fills, dashed strokes, decoded package images, transforms, clipping, and chart sub-primitives. Rectangular/path gradients and semantic placeholder primitives remain explicit magenta skips.
