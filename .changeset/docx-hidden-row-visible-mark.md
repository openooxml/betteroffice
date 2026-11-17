---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

A DOCX table row marked hidden keeps its height when any paragraph mark in it is visible, as Word does, even if the paragraphs are empty. Only rows whose paragraph marks and runs are all hidden are left out.
