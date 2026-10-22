---
"@betteroffice/xlsx": patch
---

A scalar `XLOOKUP` over a lookup or result array the formula computes now answers instead of erroring, reads each reference argument once, and returns `#VALUE!` when the lookup vector's orientation does not match the result array, as Excel does.
