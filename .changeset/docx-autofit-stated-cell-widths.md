---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/python-docx': patch
'@betteroffice/rust-crates': patch
---

Autofit tables with no table width keep their cells' stated widths when those widths and the column grid both exceed the text width, as Word does, instead of squeezing the columns and wrapping text that fits.
