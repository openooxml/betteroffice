---
'@betteroffice/vsdx-react': patch
---

Bind rows of an imported table onto a shape's data. Columns match shape-data rows by header, each bind is one undo entry through the engine's `setShapeData`, and a row the engine refuses — a date, duration or currency row, or a non-numeric value in a number row — leaves the whole bind unwritten and reports which row refused.
