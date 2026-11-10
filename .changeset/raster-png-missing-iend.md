---
'@betteroffice/pptx-native': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Native slide and page PNGs draw a PNG picture that ends without its closing chunk, as browsers do, instead of skipping it.
