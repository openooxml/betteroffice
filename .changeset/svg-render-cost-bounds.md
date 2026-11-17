---
'@betteroffice/pptx-native': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-pptx': patch
---

SVG and metafile pictures in native slide PNGs render within a bounded cost per picture and per slide, drawing at a reduced scale when a picture would cost more.
