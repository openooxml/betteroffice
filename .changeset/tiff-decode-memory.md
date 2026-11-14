---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/pptx': patch
'@betteroffice/pptx-native': patch
'@betteroffice/rust-crates': minor
---

TIFF pictures whose decoded size their dimensions don't bound, such as JPEG-compressed ones, are no longer decoded, keeping memory use predictable. Adds `tiff_bounded_pixels`. DOCX opens charge a per-document TIFF pixel budget. VSDX rendering skips TIFF pictures.
