---
'@betteroffice/pptx': patch
'@betteroffice/pptx-native': patch
'@betteroffice/rust-crates': patch
---

Pictures whose fill sits inside alternate content, as Mac PowerPoint saves a PDF picture with a PNG fallback, now draw the branch the renderer supports instead of being left out.
