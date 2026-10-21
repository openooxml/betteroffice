---
"@betteroffice/docx": patch
---

Moving the caret and typing in long docx documents resolve the caret position once per document version instead of walking the story again for every selection, sticky-position and frame lookup.
