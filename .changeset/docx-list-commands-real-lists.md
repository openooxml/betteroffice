---
'@betteroffice/docx': minor
'@betteroffice/docx-react': patch
'@betteroffice/docx-native': minor
'@betteroffice/rust-crates': minor
---

Bullets and Numbering create a list of their kind or join the one directly above or below, put a selection on one list, and lay it out as saved. The engine gains `ensure_list_definition_after` and `paragraph_style_numbers`; `ensure_list_definition` also returns `listLevelIndents`; an engine without `ensureListDefinition` refuses `toggleList`.
