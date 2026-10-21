---
"@betteroffice/docx": patch
---

Typing in long docx documents no longer re-decodes the measurements of unchanged paragraphs on every page it redraws: frames reuse their converted extents.
