---
'@betteroffice/rust-crates': minor
---

Adds `docx_edit::layout_request::region_layout_request`, which builds the editor's region layout request from a parsed DOCX package. DOCX `render_png` now uses it, so theme colours, paragraph styles and section settings render as in the editor.
