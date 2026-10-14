---
"@betteroffice/docx-native": patch
"@betteroffice/pptx-native": patch
"@betteroffice/python-docx": patch
"@betteroffice/python-pptx": patch
"@betteroffice/rust-crates": patch
---

Native page and slide rendering holds each image's decoder to the image budget left, so a GIF whose frame is larger than its declared size is skipped instead of allocating past the budget.
