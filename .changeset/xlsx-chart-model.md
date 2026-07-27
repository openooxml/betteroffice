---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Classic charts are now part of the workbook model rather than an opaque preserved part. Their series, category and title references follow row and column edits, sheet renames and sheet removals the same way cell formulas do, the cached points beside each reference are regenerated from the edited workbook, anchors move and resize according to their `editAs` mode, and the chart part is patched in place so unmodelled chart markup survives. Structural edits on a workbook whose charts are all covered are no longer refused; a ChartEx part, an unclaimed chart part, a pivot-sourced or externally cached chart, a `sqref` extension reference, and a cache beside a reference that is not a direct one-dimensional range still refuse them. Charts are preserved from the package they were read with; creating one is not supported, and a chart-bearing model with no source package is refused rather than saved without its charts.
