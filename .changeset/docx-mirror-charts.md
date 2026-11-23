---
"@betteroffice/docx": patch
"@betteroffice/docx-native": patch
"@betteroffice/rust-crates": minor
---

Screen reader page mirrors read a chart as a figure holding its values as a table, series as column headers and categories as row headers, formatted like its data labels. A chart without cached values says it has no data. `ChartA11yAttrs` literals must set `data`.
