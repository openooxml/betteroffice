---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

A footnote or endnote reference lost its revision when its paragraph was edited: a deleted reference came back as live. It now keeps its revision, formatting applied beside it is saved, and one suggested for deletion inside a field result stays in the file.
