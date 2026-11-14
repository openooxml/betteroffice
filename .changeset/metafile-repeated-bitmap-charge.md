---
'@betteroffice/docx-native': patch
'@betteroffice/pptx-native': patch
'@betteroffice/rust-crates': patch
---

EMF and WMF pictures that draw the same bitmap several times now rasterize within the image budget in native PNG exports: a repeated bitmap's decoded pixels are counted once, as the renderer decodes and frees each draw on its own.
