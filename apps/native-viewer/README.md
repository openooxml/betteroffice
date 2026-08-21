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

The fixed top toolbar provides undo, redo, bold, italic, underline, paragraph alignment, zoom, and save controls. A fixed status line shows the file and current page, sheet, or slide, plus transient save results. XLSX and PPTX keep the editing controls visible but disabled. Scroll vertically with the trackpad or mouse wheel. Hold Command or Control while scrolling to zoom. The read-only formats also accept `+`, `-`, and `0`; use Command or Control with those keys while editing DOCX.

DOCX documents are editable in the window. Click or drag to place a caret or select text, Shift-click to extend a selection, and double-click to select a word. Arrow keys, Home, End, typing, Backspace, Delete, and Enter use the resident DOCX engine. Command-Z or Control-Z undoes and adding Shift redoes. Press Command-S or Control-S to save beside the input as `<name>-edited.docx`; the viewer reopens the saved file immediately. Save results and fidelity refusals appear in the status line. XLSX and PPTX remain read-only.

DOCX saving round-trips text edits, bold, italic, underline, and alignment in supported top-level body paragraphs. The save patches only edited paragraph XML into the source package, so untouched paragraphs and unrelated package parts remain byte-for-byte unchanged. Saving refuses paragraph splits or merges, nested table or content-control paragraph edits, combined text-and-character-format changes it cannot project safely, and text or character-format edits to paragraphs containing hyperlinks, comment or note references, revision marks, drawings or embedded objects, content controls, bookmarks, fields, math, or other non-text run content. The refusal identifies the paragraph and the markup at risk before writing an output file.

The DOCX translation covers positioned glyph runs, fallback text shaping, rectangles, line and shape paths, embedded data-URL and scoped relationship images with crop/flip/rotation, and decorations. Advanced DrawingML effects and paint, filtered or bordered images, secondary-color lines, and compound or wave borders remain explicit skips.

The XLSX translation covers fills, clipped solid/dashed/dotted/double lines, geometry paths, and Carlito-shaped text with alignment, synthetic bold/italic, highlight, underline, strike, and dashed underline. Font-family requests intentionally follow `xlsx-raster` and use its bundled Carlito fallback.

The PPTX translation uses the slide layout engine and replays its positioned glyphs with the registered Arial-compatible faces. It covers positioned glyph runs, underline, normalized shape paths, solid and linear/radial gradient fills, dashed strokes, decoded package images, transforms, clipping, and chart sub-primitives. Rectangular/path gradients and semantic placeholder primitives remain explicit magenta skips.
