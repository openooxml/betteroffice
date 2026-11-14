---
'@betteroffice/rust-crates': patch
---

`Document::render_png` on an unedited document now lays out from the open model and builds only the requested page instead of saving and reopening first. The PNG is unchanged; an edited document still renders through a save.
