---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/docx': patch
'@betteroffice/pptx': patch
---

Number format codes in styles and chart caches are read and written as `ST_Xstring`, as Excel does: its `_x000a_` line break reads as a line break, any format set through the API saves as valid XML, and a literal `_xHHHH_` round-trips.
