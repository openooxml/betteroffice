---
"@betteroffice/xlsx": patch
---

Formulas that union references in parentheses, such as `RANK(A1,(B1:B9,D1:D9))` or `SUM((A1:A3,C1))`, now recalculate, and `AREAS` and `RANK.AVG` are supported.
