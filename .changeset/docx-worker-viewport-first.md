---
'@betteroffice/docx': patch
'@betteroffice/docx-react': patch
---

The worker-only editor paints a long document's first pages once they lay out exactly with their fonts, including for `previewFirstPage`, and finishes the rest in the background. Documents with page totals, notes, floats or columns open fully laid out.
