---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/python-docx': patch
---

A DOCX list level override that sets no start value now keeps the start of the level it overrides, so its first label reads "A." or "3." instead of nothing or "0.".
