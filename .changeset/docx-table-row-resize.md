---
'@betteroffice/docx': patch
'@betteroffice/docx-react': patch
'@betteroffice/docx-native': patch
---

Table row borders can be dragged to set a row's height in both docx editors. The height saves as an at-least `w:trHeight`, and a drag is one undo step. Adds `set_row_height` to the engine.
