---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Charts are now part of the workbook model rather than an opaque preserved part. Their series, category and title references follow row and column edits, sheet renames and sheet removals the same way cell formulas do, anchors move and resize according to their `editAs` mode, and the chart part is patched in place so unmodelled chart markup survives. Structural edits on a workbook containing charts are no longer refused.
