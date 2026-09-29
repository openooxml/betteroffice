---
"@betteroffice/docx": minor
"@betteroffice/rust-crates": minor
---

The DOCX editor now draws EMF, EMF+ and WMF pictures instead of leaving them blank; one it cannot draw shows a neutral placeholder and is reported in `document.warnings`, and saving keeps the original bytes. The new `betteroffice-metafile` crate holds the replay, enabled in `betteroffice-docx-parse` and `betteroffice-docx-edit` by an opt-in `metafile` feature.
