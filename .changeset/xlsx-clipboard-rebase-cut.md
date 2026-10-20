---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
---

Copied formulas now rebase relative references by default; `CellInputEdit.copyFrom` opts batch edits into rebasing. Pressing Escape before a queued cut runs cancels it. Array and spill copies are refused; `CellEdit.isArrayFormula` identifies their cells. Oversized clipboard pastes are refused.
