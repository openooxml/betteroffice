---
"@betteroffice/docx": patch
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Slide and page image decoding now holds each image to its remaining budget, so a GIF whose frame is larger than its declared size is skipped instead of allocating past the budget.
