---
"@betteroffice/xlsx": patch
---

`=ISOWEEKNUM()` with no argument or more than one now evaluates to `#VALUE!` instead of crashing the calculation engine, which in the browser ended the document worker.
