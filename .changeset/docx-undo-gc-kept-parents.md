---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Collaborative docx editing no longer crashes when a peer deletes a story or table this replica edited and undid. Once undo history attaches, a replica keeps deleted content until reload, so undo can always restore it. Raw yrs undo bridges should call `EditingDoc::keep_deleted` first.
