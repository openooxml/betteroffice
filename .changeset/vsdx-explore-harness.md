---
"@betteroffice/rust-crates": patch
---

Add a private-corpus survey binary to the vsdx-bench crate, behind a `VSDX_EXPLORE_DIR`
environment variable, reporting parse, round-trip, resolve, evaluate, render, geometry-row and
section-visibility counts per file and in aggregate.

Histogram keys are reduced to bounded categories so no document content reaches the output.
