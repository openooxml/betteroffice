---
"@betteroffice/docx": patch
---

Saving a document that inserts or removes paragraphs or table rows keeps the untouched XML of parts whose tables have vertically merged cells, instead of rewriting the whole part.
