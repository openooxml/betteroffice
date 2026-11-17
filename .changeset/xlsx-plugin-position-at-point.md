---
'@betteroffice/xlsx-react': patch
---

Spreadsheet plugins can hit-test the grid: `geometry.getPositionAtPoint(clientX, clientY)` now answers the cell under a point, matching the editor's own answer, in every editor. It returns null outside the grid, over charts and after the workbook is replaced.
