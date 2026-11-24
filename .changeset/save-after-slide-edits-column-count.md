---
'@betteroffice/rust-crates': patch
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/python-docx': patch
'@betteroffice/python-pptx': patch
---

`Presentation::save` in Rust and `Presentation.save` in Python succeed after slides are deleted and prepare fonts for added slides. DOCX layout limits a section's column count to what Word supports, so a very large count no longer stalls `Document::layout`.
