---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/python-docx': patch
---

DOCX numbered list levels without a start value now count from zero, as Word does, and a zero counter prints "0." instead of nothing.
