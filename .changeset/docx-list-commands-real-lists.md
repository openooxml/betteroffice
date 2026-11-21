---
'@betteroffice/docx': minor
'@betteroffice/docx-react': patch
'@betteroffice/docx-native': minor
'@betteroffice/rust-crates': minor
---

Bullets and Numbering now create a list of their kind or join the one directly above or below, put a selection on one list, and lay it out as it is saved. The engine gains `ensure_list_definition_after` and `paragraph_style_numbers`; an engine without `ensureListDefinition` refuses `toggleList`.
