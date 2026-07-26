# betteroffice-drawingml

The DrawingML models shared by the DOCX, XLSX, and PPTX engines, so the three
formats resolve graphics identically.

- `color` — the DrawingML color model and its resolution against a theme,
  including the transform stack (tint, shade, alpha, luminance modulation)
- `theme` — theme color schemes, font schemes, and format schemes
- `geometry` — the preset shape geometries and the custom-geometry path model
- `shape` — shape properties: outlines, fills, and effects
- `picture` — picture fills and image references
- `chart` — chart parsing and plot geometry, behind the `chart` feature

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
