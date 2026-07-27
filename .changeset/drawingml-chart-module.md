---
"@betteroffice/rust-crates": patch
---

Chart parsing and plot geometry now live in `betteroffice-drawingml` behind a `chart` cargo feature, shared by the docx, xlsx and pptx engines instead of being duplicated per format.

Breaking, for Rust consumers of `betteroffice-drawingml` only. `ChartAxes`, `ChartAxis`, `ChartLegend`, `ChartMarker`, `ChartPoint`, `ChartPlotGroup` and `ChartSeries` were re-exported at the crate root in 0.0.4 and are no longer. Enable the `chart` feature and import them from `ooxml_drawingml::chart`. The feature is off by default, so a crate that does not draw charts no longer compiles the geometry engine.
