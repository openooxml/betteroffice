---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

Printed sheets keep their stored row heights when the default row height is below the font's natural height, so they paginate closer to Excel.
