---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Rendering a sheet without an explicit range no longer fails on a workbook whose chart is parked far below the data. The frame still grows to take in a chart near the used range, but one that would push the image past the renderer's pixel caps is now left out of frame — the way a chart past the edge of an explicit viewport already is — instead of taking the whole render down with it. A render capped by `max_width` / `max_height` that used to pad out to the cap around such a chart now returns the tighter used-range frame, so its reported dimensions can be smaller than before.
