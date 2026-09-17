---
'@betteroffice/docx': patch
---

Apply the document grid to line heights: sections with an activating `w:docGrid` type (`lines`, `linesAndChars`, `snapToChars`) snap each line's height up to the next grid-pitch multiple, honouring paragraph/run `w:snapToGrid` opt-outs. Grids with `default` type or a bare `linePitch` never snap.
