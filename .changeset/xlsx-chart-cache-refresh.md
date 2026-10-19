---
"@betteroffice/xlsx": patch
---

Edited saves refresh the cached values of chart series whose source cells changed, so viewers that draw from the cache show the edited data; untouched charts keep their bytes.
