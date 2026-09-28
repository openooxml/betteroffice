---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
"@betteroffice/python-docx": patch
---

Lay out DOCX pages closer to Word: exact-height table rows stay whole across pages, a table's leading edge follows the document's `compatibilityMode`, inline images keep their declared extent, a paragraph that anchors a shape stays on one page, full-width wrap bands and floats in table cells or around floating tables reserve their space, text wraps on both sides of an interior float, and space-before is kept after an inline or paragraph page break. Lines in East Asian, substituted Latin and substituted Korean faces are measured as Word measures them, a Latin line takes its face's `hhea` span, and a tab past the stop the pen rests on moves to the next stop.
