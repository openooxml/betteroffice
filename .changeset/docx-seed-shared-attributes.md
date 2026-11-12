---
'@betteroffice/rust-crates': patch
---

When a DOCX is first loaded for editing, equal attribute keys and values share one allocation. Retained memory after opening a 5 MiB DOCX drops by 41.30 MiB native and 27.69 MiB wasm. Peak memory and saved bytes are unchanged.
