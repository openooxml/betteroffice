---
"@betteroffice/xlsx": patch
---

Formulas the recalculation budget cuts short keep their saved values under `IFERROR`, `ISERROR` and the other `IS` functions instead of showing the fallback, and show `#NUM!` when there is no saved value.
