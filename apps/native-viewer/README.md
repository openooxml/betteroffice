# BetterOffice native viewer

This experimental macOS-first app lays out a DOCX with BetterOffice and paints its native display list with Vello.

From this directory, render and compare the first page:

```sh
cargo run --release -- --png page.png
```

The command writes `page.png` through Vello and `page.raster.png` through the existing raster backend, then prints image-difference metrics.

Use `--document FILE`, `--page N`, and `--scale N` to select another DOCX, one-based page, or output scale. Every display-list primitive that cannot be translated is replaced by a magenta box and cross, counted by type, and reported with its reason.

Open the interactive viewer:

```sh
cargo run --release
```

Scroll vertically with the trackpad or mouse wheel. Hold Command or Control while scrolling to zoom, or use `+`, `-`, and `0`.

The current translation covers positioned glyph runs, fallback text shaping, rectangles, line and shape paths, scoped relationship images with crop/flip/rotation, and decorations. Advanced DrawingML effects and paint, filtered or framed images, secondary-color lines, and compound or wave borders remain explicit skips.
