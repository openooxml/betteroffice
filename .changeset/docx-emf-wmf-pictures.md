---
"@betteroffice/docx": minor
"@betteroffice/rust-crates": minor
---

Draw EMF, EMF+ and WMF pictures in the DOCX editor. They used to paint nothing. Parsing now replays each metafile's records into an SVG display copy: lines, polygons, rectangles, ellipses, arcs and Béziers with their pens, dashes and brushes, text runs positioned per character in their fonts, clipping, world and page transforms, and embedded bitmaps. A dual EMF+ metafile draws from its EMF+ records, as Office does, and falls back to its GDI records when those hold something the replay cannot draw. A metafile the replay refuses shows a neutral placeholder, and ink it drew without is listed; both cases are reported in `document.warnings`. The canvas also paints that placeholder for any embedded picture the browser cannot decode. Saving keeps the original metafile bytes.

The new `betteroffice-metafile` crate holds the bounded replay: `replay` for the drawing, `to_svg` behind the `svg` feature, and the strict vector `decode` that PowerPoint pictures already used, moved here from `betteroffice-pptx-render` unchanged. `betteroffice-docx-parse` and `betteroffice-docx-edit` convert pictures behind a new `metafile` feature, and `betteroffice-drawingml` exposes its PNG writer behind a `png` feature.
