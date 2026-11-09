---
'@betteroffice/pptx-native': patch
'@betteroffice/rust-crates': patch
---

Slides with a CMYK TIFF picture that has an alpha channel now export on the native engine instead of failing with `picture-unavailable`, using the same decoder the canvas engine uses.
