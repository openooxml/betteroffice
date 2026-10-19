---
'@betteroffice/xlsx-react': patch
'@betteroffice/xlsx': patch
---

The xlsx editor keeps the active cell at the start of a range selection, as Excel does: after a drag, Shift+click or Shift+arrow the name box, formula bar and typing target stay on the anchor cell, and plain arrows, Tab and Enter move from it. Redoing a cell edit selects the edited cell.
