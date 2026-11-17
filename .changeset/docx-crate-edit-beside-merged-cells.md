---
'@betteroffice/rust-crates': patch
'@betteroffice/python-docx': patch
---

Editing a paragraph in a Word document whose tables have vertically merged cells no longer rewrites the whole document body when saved through the Rust crate or the Python package; the bytes outside the edit are kept.
