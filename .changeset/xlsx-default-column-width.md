---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

On screen, XLSX columns without their own width now use the sheet's `defaultColWidth`, or its `baseColWidth` when that is absent, instead of always 64px.
