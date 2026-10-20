---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Collaborative editing no longer crashes when a peer deletes a story or table this replica edited and undid, because a replica with undo history keeps deleted content until reload. Raw yrs undo bridges should call `EditingDoc::keep_deleted` first.
