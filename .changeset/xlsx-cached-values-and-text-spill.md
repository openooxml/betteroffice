---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
---

Keep the value a workbook already carries when recalculation names a function the engine does not implement, so a sheet built on `OFFSET`, `MMULT` or another gap renders the numbers its authoring app computed instead of `#NAME?`; a formula with no cached value still reports the error, and a supported function still overwrites a stale cache. Text that outgrows its cell now spills over the blank cells beside it, towards whichever side its alignment points, stopping at the first neighbour that draws something, at a merge, and at the edge of the laid-out columns; wrapped, shrink-to-fit, merged and non-text cells stay boxed.
