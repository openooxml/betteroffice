---
'@betteroffice/agents': patch
'@betteroffice/docx': patch
---

`insert_paragraphs` and `set_paragraph_style` now apply paragraph styles that define list numbering, such as List Bullet or List Number, so the paragraphs become real list items instead of being refused.
