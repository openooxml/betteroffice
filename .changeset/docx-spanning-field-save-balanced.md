---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Editing a paragraph of a field that spans paragraphs, such as the last entry of a table of contents, no longer saves the field without its end or with an extra one. Such a save writes the whole part instead.
