---
'@betteroffice/rust-crates': patch
---

Adds `docx_edit::layout_request::measurement`. `betteroffice-docx` `render_png` and `Document::layout` now measure text with the document's compatibility settings (compatibility mode, `noLeading`, `doNotExpandShiftReturn`) and authoritative shaping, as the editors do.
