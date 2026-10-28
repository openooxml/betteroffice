---
"@betteroffice/xlsx": patch
---

`SUBTOTAL` and `AGGREGATE` now skip other subtotals inside their ranges, so grand totals over grouped data no longer count the group subtotals twice.
