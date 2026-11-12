---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': minor
---

Number formats now round displayed digits as Excel does, from the value at 15 significant digits: `0.00` shows 1.005 as 1.01 and `0.00%` shows 0.00035 as 0.04%, and fraction parts follow, in cells and in `TEXT`.
