---
'@betteroffice/rust-crates': patch
'@betteroffice/python-docx': patch
'@betteroffice/python-pptx': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/python-vsdx': patch
---

Adds opt-in `decode_font_face` calls for zstd and woff2t font pack faces. Rust facades expose the shared `betteroffice-fonts` decoder as `fonts`; Python returns font bytes and an optional transformed payload.
