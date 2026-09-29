---
"@betteroffice/docx": minor
"@betteroffice/rust-crates": minor
---

The DOCX editor now draws EMF, EMF+ and WMF pictures; one it cannot draw shows a placeholder and a `document.warnings` entry. The new `betteroffice-metafile` crate does the replay, enabled in `betteroffice-docx-parse` and `betteroffice-docx-edit` by an opt-in `metafile` feature.
