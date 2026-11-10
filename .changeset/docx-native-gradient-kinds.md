---
'@betteroffice/docx-native': patch
'@betteroffice/python-docx': patch
'@betteroffice/rust-crates': patch
---

Native DOCX PNG export paints rectangular and path gradient fills as a radial gradient, as the browser renderer does, instead of failing the page.
