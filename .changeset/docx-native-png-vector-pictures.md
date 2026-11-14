---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/pptx': patch
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-docx': patch
---

Native docx page PNGs draw SVG, EMF and WMF pictures. A picture the renderer still cannot draw rejects `renderPng` with the new `picture-unavailable` code instead of leaving it blank.
