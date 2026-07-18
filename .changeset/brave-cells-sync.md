---
"@betteroffice/rust-crates": patch
"@betteroffice/xlsx": patch
---

Add deterministic Yrs replicas, bounded and validated v1 update exchange, and
nonstructural cell and dimension collaboration to the Rust XLSX workbook API.
Collaborative inverse-op undo and redo remain disabled until a Yrs-aware undo
manager can preserve concurrent edits.
