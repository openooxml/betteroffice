---
'@betteroffice/docx': minor
'@betteroffice/docx-react': minor
'@betteroffice/rust-crates': minor
---

Reads now include `drawing` (`wp:docPr` properties, placement and type) on drawing atoms by default, as do structured image and unsupported inlines. Plugins request a drawing's frame rects with the new `drawing` geometry target. Sessions seeded earlier read `drawing` as absent until reseeded. Supply `drawing` on `TextAtom` and image/unsupported `InlineKind` and `DocxExportInline` values, `drawing_properties` on `RunContent::Drawing`, `RunContent::Shape`, `RunContent::Chart` and `RunContent::OpaqueDrawing`, and `ChartA11yAttrs.frame` in literals.
