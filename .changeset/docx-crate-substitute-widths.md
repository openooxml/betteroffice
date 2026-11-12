---
"@betteroffice/rust-crates": patch
"@betteroffice/python-docx": patch
---

The `betteroffice-docx` crate measures a substitute face with the requested family's advance widths and line metrics, as the editor and the wasm build do, so its page counts for documents naming families without a bundled face match theirs more often.
