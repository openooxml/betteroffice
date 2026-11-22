---
'@betteroffice/rust-crates': patch
'@betteroffice/pptx': patch
'@betteroffice/python-pptx': patch
---

Saving an edited slide with many identical shapes no longer takes time that grows with the square of their count. The splice that keeps untouched XML now has a work bound and regenerates the slide past it.
