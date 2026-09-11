---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Stroke a chart series' line at the width its own `c:ser/c:spPr/a:ln` declares instead of a fixed 2px, and draw no line at all when that outline is `a:noFill`.

Migrate collaboration snapshots to schema 21, importing series lines from a reattached source.
