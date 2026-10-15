---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Fixes clipped currency ticks and bar category labels in xlsx charts. Adds provided `PlotSink::anchors_text_end` and `PlotSink::push_end_aligned_text` methods so a sink can anchor text at its right edge.
