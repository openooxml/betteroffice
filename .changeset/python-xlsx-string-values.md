---
'@betteroffice/python-xlsx': patch
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
---

Honor Text cell formatting when entering values, including numeric strings, booleans, and formula-looking text. General cells retain Excel-style input coercion. Apply the same format-aware behavior to cell assignments, batches, and proposals.
