# BetterOffice native viewer

This experimental macOS-first app paints BetterOffice DOCX and XLSX display lists with Vello.

From this directory, render and compare the first DOCX page:

```sh
cargo run --release -- --png page.png
```

Render the used range of the first sheet in the XLSX showcase:

```sh
cargo run --release -- --document ../demo/public/showcase.xlsx --sheet 1 --png sheet.png
```

Each command writes its requested PNG through Vello and a sibling `.raster.png` through the matching existing raster backend, then prints image-difference metrics.

Use `--document FILE` and `--scale N` for either format. `--page N` selects a one-based DOCX page, while `--sheet N` selects a one-based XLSX sheet. XLSX defaults to the selected sheet's used range and resolves supported charts from the workbook package. Every display-list item that cannot be translated is replaced by a magenta box and cross, counted by type, and reported with its reason.

Open the interactive viewer:

```sh
cargo run --release
```

Scroll vertically with the trackpad or mouse wheel. Hold Command or Control while scrolling to zoom, or use `+`, `-`, and `0`.

The DOCX translation covers positioned glyph runs, fallback text shaping, rectangles, line and shape paths, scoped relationship images with crop/flip/rotation, and decorations. Advanced DrawingML effects and paint, filtered or framed images, secondary-color lines, and compound or wave borders remain explicit skips.

The XLSX translation covers fills, clipped solid/dashed/dotted/double lines, geometry paths, and Carlito-shaped text with alignment, synthetic bold/italic, highlight, underline, strike, and dashed underline. Font-family requests intentionally follow `xlsx-raster` and use its bundled Carlito fallback.
