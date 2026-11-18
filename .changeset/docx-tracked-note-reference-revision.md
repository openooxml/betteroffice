---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
---

A footnote or endnote reference lost its revision and its character style when its paragraph was edited: a deleted reference came back as live and unstyled. It now keeps both, and one suggested for deletion inside a field result stays in the file.
