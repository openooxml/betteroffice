---
'@betteroffice/pptx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-pptx': patch
---

Honour `a:normAutofit` as PowerPoint renders it: the stored `fontScale` (0.1–1.0) scales sizes, rounded to whole points below 1.0, and `lnSpcReduction` (0–0.9) is subtracted from percentage line spacing, including the implicit single-spaced default, so shrink-to-fit bodies keep PowerPoint's font size and line pitch.
