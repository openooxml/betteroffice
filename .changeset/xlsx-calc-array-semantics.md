---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: recalculation now reads an argument that computes an array wherever a range is accepted, resolves whole-row and reference-range references, and answers per element where Excel does. Together with twenty-three functions that had no implementation — `DATEVALUE`, `FV`, `TREND`, `FREQUENCY`, `YEARFRAC`, `NORM.DIST`, `ADDRESS`, `HYPERLINK`, the `.INTL` workday pair and the `-A` aggregates among them — this covers the idioms real workbooks are built from: `MEDIAN(IF(range=key,range))` and the other CSE aggregates, `INDEX(data,MATCH(key,a&b,0))` over concatenated keys, `AGGREGATE(15,6,ROW(range)/(range=key),n)` for the nth match, `SUMPRODUCT(--(range=key))`, `ROWS($1:1)` as a running counter, and `AA9:INDEX(AA9:AH9,n)` as a computed endpoint.

`INDEX` now picks its reference or array form from its first argument as Excel does, `MATCH` answers once per key, and a formula whose result is an empty reference stores 0 rather than staying blank. A whole-column reference costs the rows the sheet actually reaches, so a handful of `COUNTIFS(F:F,…)` formulas no longer exhaust the recalculation budget and leave the rest of the workbook reporting `#NUM!`.

On the 165-workbook calculation benchmark, which strips cached values and compares against Excel's own full rebuild, this moves the corpus from 92.0076% to 99.5993% of cells matching Excel — 133 of 165 workbooks now match exactly, up from 49 — with no workbook regressing. LibreOffice scores 85.8713% on the same corpus.
