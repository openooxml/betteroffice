---
'@betteroffice/pptx-react': patch
---

The PowerPoint editor keeps its private crash-recovery copy of the opened file as a Blob instead of in page memory, lowering page memory by the file's size.
