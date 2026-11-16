---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

XLSX rows keep their authored heights when a sheet declares an unusual default row height. Wrapped text matches Excel more closely: Arial-like fonts use their own line spacing, and text taller than its row shows only the lines that fit.
