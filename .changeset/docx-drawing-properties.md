---
'@betteroffice/docx': minor
'@betteroffice/docx-react': minor
'@betteroffice/rust-crates': minor
---

Drawing atoms and structured image/unsupported inlines expose `drawing` (`wp:docPr` properties, placement and type) by default. Plugins request drawing frame rects via the `drawing` target. Previously seeded sessions lack `drawing` until reseeded. Set new `TextAtom`, image/unsupported `InlineKind` and `DocxExportInline`, `RunContent::Drawing`, `RunContent::Shape`, `RunContent::Chart`, `RunContent::OpaqueDrawing` and `ChartA11yAttrs` fields.
