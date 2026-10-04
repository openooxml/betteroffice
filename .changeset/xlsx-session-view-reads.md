---
"@betteroffice/xlsx": minor
---

Experimental workbook sessions add `call.sheetView`, `call.cellGeometry` and `call.cellInputs` for worker-owned viewing, and `openWorkbookSession` accepts a `signal` that cancels an open in flight. Session frames include visible `mergedRanges` by default; session use remains opt-in.
