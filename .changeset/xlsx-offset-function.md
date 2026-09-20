---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Add `OFFSET` to the formula engine. It returns a reference, not a value, so it feeds every area-taking function — `SUM`, `AVERAGE`, `VLOOKUP`, `MATCH`, `INDEX` and a nested `OFFSET` all accept one. `height` and `width` default to the anchor's own extent and may be negative, extending the rectangle back from the shifted corner; a zero size or a rectangle that leaves the sheet is `#REF!`, and a multi-cell result used as a scalar is `#VALUE!`, matching the engine's existing no-implicit-intersection rule for bare ranges. The dependency graph carries the exact resolved rectangle whenever the offsets and sizes are literal, so a downstream cell recalculates when its real target is edited; where an offset is computed the target cannot be known before evaluation, so the calling cell is marked volatile rather than left reading stale values.
