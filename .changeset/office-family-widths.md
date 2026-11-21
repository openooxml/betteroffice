---
"@betteroffice/fonts": patch
"@betteroffice/docx": patch
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": minor
---

Substituted Office fonts, including Aptos, Arial Narrow, Cambria and Verdana, measure at the real font's glyph widths and line heights, so text breaks like Word and PowerPoint. Legacy names such as Tms Rmn and Helv also resolve. Set `substitute_scale` in `TextRunPrimitive` literals.
