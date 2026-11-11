---
'@betteroffice/rust-crates': patch
---

Adds `docx_edit::layout_request::measurement`. `betteroffice-docx` now measures text with the document's compatibility settings (compatibility mode, `noLeading`, `doNotExpandShiftReturn`) and authoritative shaping, as the editors do.
