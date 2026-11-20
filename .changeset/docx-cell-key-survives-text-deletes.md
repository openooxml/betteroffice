---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Backspace, deleting selected text and IME composition inside a table cell no longer rebuild the cell location cache on every keystroke. The cache now changes only when tables, cells, embeds or paragraph marks may have changed.
