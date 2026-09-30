---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Pages stay painted, and scroll-to calls land on target, when a host scales the editor with CSS `zoom` on an ancestor. New `effectiveZoom(element)` export from `@betteroffice/docx/layout/render`.
