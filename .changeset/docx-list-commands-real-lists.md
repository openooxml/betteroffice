---
'@betteroffice/docx': minor
'@betteroffice/docx-react': patch
'@betteroffice/docx-native': minor
'@betteroffice/rust-crates': minor
---

Bullets and Numbering now create a list of their kind or join the one directly above or below, put a selection on one list, and lay it out as it is saved. `ensureListDefinition` takes an optional first paragraph; an engine without it refuses `toggleList`.
