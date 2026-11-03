---
'@betteroffice/docx': patch
'@betteroffice/fonts': patch
---

Opening a document no longer copies a whole bundled font just to check which characters it covers, which lowers peak memory while opening CJK documents.
