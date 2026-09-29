---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Ship a page whose positions all moved by one delta as a single position-shift run. A frame-delta shift run can now carry a present-only flag that shifts just the masked fields each primitive has, so a page of text, table rects, widgets and position-free primitives moves with one run instead of one run per primitive. A note area whose only change is a moved reference anchor now ships as a shift too, carrying the new anchors, instead of redrawing its page. An edit near the start of a long document sends a few kilobytes of shifts instead of megabytes, and the main thread decodes them and updates the query store accordingly.
