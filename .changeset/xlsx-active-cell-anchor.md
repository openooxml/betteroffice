---
'@betteroffice/xlsx-react': patch
'@betteroffice/xlsx': patch
---

The xlsx editor keeps the active cell on the anchor of a range selection, as Excel does: the name box, formula bar and typing stay there after drag, Shift+click or Shift+arrow, and arrows, Tab and Enter move from it. Redo selects the edited cell.
