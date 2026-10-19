---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
---

Copied formulas now rebase relative references by default; `CellInputEdit.copyFrom` opts batch edits into rebasing. Cut defers one undoable move until paste. Array and spill copies are refused; `CellEdit.isArrayFormula` identifies their cells. Oversized clipboard pastes are refused.
