---
'@betteroffice/xlsx': patch
---

`ROUND` now matches Excel on values that sit at a decimal tie or carry more than 15 significant digits: `ROUND(1.005,2)` is 1.01 and `ROUND(1000000000000000+0.75,0)` is 1000000000000000.
