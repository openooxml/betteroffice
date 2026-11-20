---
'@betteroffice/docx-react': patch
'@betteroffice/docx': minor
'@betteroffice/docx-native': minor
'@betteroffice/rust-crates': minor
'@betteroffice/xlsx': minor
'@betteroffice/pptx': minor
---

A DOCX document opened with a seed and without collaboration recovers after a worker restart following undo or redo; undo history starts afresh. Adds Rust `apply_verbatim_update`, `clientId`, `onStream`, an `applyUpdate` command kind, and optional stream members: the host's `recordUpdate`, `takeStream`, `withCommandWork`, and the engine's `applyVerbatimUpdate`.
