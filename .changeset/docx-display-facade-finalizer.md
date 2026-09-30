---
"@betteroffice/docx": patch
---

An unmounted or replaced `DocxEditor` no longer stays in memory with its rendered pages, which could happen when it unmounted before its latest layout was queried.
