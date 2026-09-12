---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Charts now survive a save. The drawing that places a chart, and any drawing the parser does not model, is kept as its original markup and written back verbatim, so the chart part and its relationship stay referenced; the editor save path projects chart runs back from the session instead of dropping them.
