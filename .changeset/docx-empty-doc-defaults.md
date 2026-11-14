---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/python-docx': patch
'@betteroffice/rust-crates': patch
---

Documents whose styles declare empty document defaults lay out their unstyled paragraphs and table cells without paragraph spacing, as Word does, instead of with Word's 8 pt after and 1.08 line spacing.
