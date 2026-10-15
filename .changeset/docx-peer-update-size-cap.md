---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Collaboration updates from other peers larger than 64 MiB (`MAX_UPDATE_BYTES`) are now refused before decoding, matching the xlsx and pptx limit. Loading a document state with `loadState` is unchanged.
