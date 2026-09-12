---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Parse `ppt/tableStyles.xml` and resolve a cell through the table style cascade: `wholeTbl`, the row band, `firstCol`/`lastCol`, `firstRow`/`lastRow`, then the cell's own `a:tcPr`, with each part's borders applied against the sides of its own region and an `a:noFill` edge clearing what a lower part set. A reattached source restores the styles a stored package never carried.
