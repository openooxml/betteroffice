---
'@betteroffice/docx': minor
'@betteroffice/docx-react': minor
---

Reads now include `drawing` (`wp:docPr` properties, placement and type) on drawing atoms by default, as do structured image and unsupported inlines. Plugins request a drawing's frame rects with the new `{ kind: 'drawing', paragraph, offset }` geometry target. Sessions seeded earlier read `drawing` as absent until reseeded.
