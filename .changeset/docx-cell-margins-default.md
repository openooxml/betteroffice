---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/python-docx': patch
'@betteroffice/rust-crates': patch
---

Table cells with no cell margins in the document, such as tables inserted or pasted in the editor, lay out with Word's default 0.19 cm left and right margins, so text wraps on screen where Word and LibreOffice wrap it.
