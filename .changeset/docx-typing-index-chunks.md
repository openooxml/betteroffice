---
"@betteroffice/docx": patch
---

Typing and deleting text in long docx documents no longer re-reads the whole story for every key: the edit reads the formatting around the caret from the paragraph index instead.
