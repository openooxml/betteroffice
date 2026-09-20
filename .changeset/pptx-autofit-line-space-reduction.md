---
'@betteroffice/pptx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-pptx': patch
---

Honour `a:normAutofit` the way PowerPoint renders it: the stored `fontScale` is applied verbatim and `lnSpcReduction` is subtracted from percentage line spacing — including the implicit single-spaced default — so shrink-to-fit bodies keep PowerPoint's font size and line pitch instead of being re-fitted at render time.
