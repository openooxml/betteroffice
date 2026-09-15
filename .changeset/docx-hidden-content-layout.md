---
'@betteroffice/docx': patch
---

Vanished content is now hidden by default: hidden runs, drawings, and fully hidden paragraphs stay out of the document layout while preserving source content and edit positions. Pass the `showHiddenText` render option (including the new `showHiddenText` prop on `DocxEditor`) to reveal it with normal wrapping.
