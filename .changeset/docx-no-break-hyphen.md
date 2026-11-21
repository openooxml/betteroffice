---
"@betteroffice/docx": patch
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": minor
---

Non-breaking hyphens draw and measure with the font's hyphen glyph when the font has no dedicated one, instead of a missing-glyph box. They still never break a line. Rust `TextRunPrimitive` and `GlyphRunPrimitive` literals must set `paint_text`.
