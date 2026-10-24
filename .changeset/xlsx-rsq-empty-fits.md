---
"@betteroffice/xlsx": patch
---

Recalculation now supports `RSQ` instead of returning `#NAME?`. `SLOPE`, `INTERCEPT`, `CORREL` and `RSQ` over ranges with no numeric pairs return `#DIV/0!`, as Excel does.
