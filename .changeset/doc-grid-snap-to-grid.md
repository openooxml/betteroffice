---
'@betteroffice/docx': patch
---

Apply the document grid to line heights: sections with an activating `w:docGrid` type (`lines`, `linesAndChars`, `snapToChars`) fill each `auto`-ruled line's content box up to one grid row, so a line-spacing multiple then scales the filled grid pitch — a 1.5-spaced line on a one-row grid is 1.5 rows tall, as Word renders it. A content box already taller than one row keeps its natural height. Paragraph/run `w:snapToGrid` opt-outs are honoured. Grids with `default` type or a bare `linePitch` never snap, and pinned `exact`/`atLeast` heights never snap.
