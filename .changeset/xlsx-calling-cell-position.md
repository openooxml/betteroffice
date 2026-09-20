---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
---

Resolve `ROW()` and `COLUMN()` against the calling cell, and treat `ROW`/`COLUMN`/`ROWS`/`COLUMNS` of a reference as the positional queries they are. Recalculation now tells the evaluator which cell it is evaluating, so the referenceless forms answer that cell's own 1-based row and column — the shape `OFFSET($E$6, ROW()-ROW($BP$6), 0, 1, 48)` builds its offset from. Asking where a reference sits reads nothing from it, so such an argument no longer creates a dependency edge: a cell holding `ROW($X$1)` at `$X$1` computes instead of being reported as a cycle and zeroed, and a computed argument such as `ROW(OFFSET(A1,B1,0))` is still tracked for the cells it reads. The one-argument forms, and every context built without a calling cell, are unchanged.
