---
'@betteroffice/docx': patch
'@betteroffice/docx-react': patch
'@betteroffice/fonts': patch
---

The DOCX editor keeps one copy of each bundled font instead of several, which lowers memory while opening and editing documents, most for CJK documents.
