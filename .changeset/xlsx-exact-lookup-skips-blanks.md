---
"@betteroffice/xlsx": patch
---

Exact-match `VLOOKUP`, `HLOOKUP`, `MATCH`, `XMATCH` and `XLOOKUP` no longer match a blank cell to a key of 0, "" or FALSE, so they return what Excel does (often `#N/A`, or a later match).
