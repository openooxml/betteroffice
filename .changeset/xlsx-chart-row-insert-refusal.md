---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
---

Row and column edits that reshape a range a chart plots through an unrebuildable cache, or that would split a chart part two sheets share, are now refused up front instead of failing the next save. Cell edits under a shifted range now save.
