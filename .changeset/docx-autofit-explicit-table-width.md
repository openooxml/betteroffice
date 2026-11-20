---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/python-docx': patch
'@betteroffice/rust-crates': patch
---

An autofit table whose explicit width is wider than the text area now keeps that width instead of being narrowed to the text area, matching Word. Percentage and automatic widths are unchanged.
