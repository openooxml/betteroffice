---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
---

Sheet, defined and font names are read and written as `ST_Xstring`, as Excel does: an escape Excel wrote reads decoded, and any name set through the API, including control characters, saves as valid XML and reads back unchanged.
