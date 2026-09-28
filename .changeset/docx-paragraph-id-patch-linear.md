---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Writing persisted paragraph IDs into a saved document's XML is linear again. On documents with many paragraphs that lacked IDs, such as LibreOffice or generator output, the patch step grew quadratically since paragraph IDs started being preserved across saves.
