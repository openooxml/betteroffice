---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-docx': patch
---

Keep block content controls inside table cells when a document is saved. Writing the table again (an editor save, `repackDocx`, `saveYrsDocx` after an edit or a fill, the native or Python `Document.save()`) dropped such a control together with its text; it now keeps the control, its properties and its content, also in nested tables and around tables.
