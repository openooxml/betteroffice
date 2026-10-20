---
"@betteroffice/pptx": patch
"@betteroffice/docx": patch
---

Corrects DrawingML colors by applying `transforms` with `ColorTransform` in document order by default, using linear light for tint and shade and preserving exact percentages when saving.
