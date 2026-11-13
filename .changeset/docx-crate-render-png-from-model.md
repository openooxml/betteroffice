---
'@betteroffice/rust-crates': patch
---

`Document::render_png` on an unedited document, or one changed only through `replace_paragraph_text`, now lays out from the open model and builds only the requested page instead of saving and reopening first. The PNG is unchanged; a document changed through `model_mut` still renders through a save.
