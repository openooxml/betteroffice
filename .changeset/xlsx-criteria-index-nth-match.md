---
"@betteroffice/xlsx": patch
---

Lookup tables that list the nth match of a key with `SMALL(IF(...))` array formulas over large data sheets now calculate every row instead of stopping with #NUM! partway through.
