---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': minor
'@betteroffice/python-xlsx': patch
---

Give a cell the format its column declares. A `<col>` run's `style` is now parsed, including on a run that sets no width, which every declaration used to be dropped from, and a cell with no `s` of its own takes it. A column style that names a solid fill now paints the whole column, so a sheet tinted through its columns arrives tinted instead of white, and a column filled white covers the gridlines Excel covers there. A cell's own style still wins, and an unsized row fits its content through the column style when the cell names none. On the fidelity corpus this moves `bootstrapping-calculator` from 0.72987 to 0.91027 and `wazobia-valuation-model` from 0.87910 to 0.90443, lifting the XLSX mean from 0.77498 to 0.80069 with every other sample unchanged to five decimals. Column styles stay a render projection: nothing is written back, so an untouched sheet still re-serializes byte-identically.
