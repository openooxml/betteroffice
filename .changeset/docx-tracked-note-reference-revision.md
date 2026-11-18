---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
---

A footnote or endnote reference inside a tracked insertion or deletion lost its revision when its paragraph was edited: a deleted reference came back as live. It now keeps its revision through the save.
