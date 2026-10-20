---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
---

Copied formulas now rebase relative references by default; `CellInputEdit.copyFrom` opts batch edits into rebasing. Escape cancels queued cuts; array and spill copies and oversized pastes are refused, and `CellEdit.isArrayFormula` identifies array cells.
