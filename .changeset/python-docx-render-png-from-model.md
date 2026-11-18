---
'@betteroffice/python-docx': patch
---

`Document.render_png` on an unedited document now lays out the requested page from the open document instead of saving and reopening it first. The PNG is unchanged.
