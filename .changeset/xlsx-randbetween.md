---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
---

Add `RANDBETWEEN` to the XLSX formula engine, where it previously read `#NAME?`. Its rounding was measured against Excel over 400 draws per case: a raw `bottom > top` is `#NUM!` before any rounding, the draw spans `ceil(bottom)` through `floor(top)`, and a span that rounds away to nothing yields `ceil(bottom)` — so `RANDBETWEEN(1.2, 3.8)` draws 2 or 3, never 4. The dependency graph already treated the name as volatile, so such cells re-draw on every recalculation. The stream is seedable through `EvalContext::rand_seed` for reproducible output; unseeded, each evaluation context takes a fresh stream from a process-local counter, so sibling cells differ and a run that evaluates in the same order replays the same draws.
