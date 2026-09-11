---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Paint a chart's own `c:chartSpace` fill instead of a white ground, and stroke each axis line with its own `a:ln` colour and width, or not at all under `a:noFill`.

Migrate collaboration snapshots to schema 15 after the existing schema-14 gradient-outline migration, importing chart-space fills and axis lines from a reattached source.
