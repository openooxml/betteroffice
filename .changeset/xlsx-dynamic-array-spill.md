---
'@betteroffice/xlsx': minor
'@betteroffice/rust-crates': minor
---

Evaluate array formulas as arrays and spill them across the range they occupy. A cell stored as `<f t="array" ref="A1:G100">` now evaluates once and writes its whole rectangle instead of only its anchor, ranges and operators produce blocks, and single-argument scalar functions such as `TRIM` and `CLEAN` run once per element when they receive one. `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `CHOOSECOLS`, `CHOOSEROWS`, `HSTACK`, `VSTACK`, `TOCOL`, `TOROW`, `EXPAND`, `DROP`, `TAKE` and `ANCHORARRAY` compute, `TRANSPOSE`, `INDEX`, `MATCH`, `ROW`, `COLUMN`, `MMULT`, `SUMPRODUCT`, `IF`, `IFERROR` and the numeric aggregates understand blocks, and `MAXIFS`, `MINIFS` and `LOOKUP` are implemented. Array constants (`{1,2;3,4}`) and omitted arguments (`INDEX(a,,2)`) parse. Over 165 recalculated corpus workbooks scored against Excel's own results, cell accuracy rises from 68.49% to 84.18% and the workbooks scoring nothing at all fall from 36 to 22.
