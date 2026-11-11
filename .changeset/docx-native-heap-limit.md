---
'@betteroffice/docx-native': patch
'@betteroffice/docx': patch
'@betteroffice/docx-react': patch
'@betteroffice/rust-crates': patch
---

The native DOCX engine now enforces the worker heap limit: a call past it fails with the same out-of-memory error as wasm, and its object refuses later calls. Heap counting, and its cost, applies only while a limit is set.
