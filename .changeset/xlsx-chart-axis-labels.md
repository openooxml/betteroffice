---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Prevents clipped currency ticks and bar category labels in xlsx charts by default. Adds opt-in `PlotSink::anchors_text_end` and `PlotSink::push_end_aligned_text` capabilities for anchoring text at its right edge. Hidden secondary axes retain the base plot reserve.
