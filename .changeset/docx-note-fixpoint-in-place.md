---
"@betteroffice/rust-crates": patch
---

Replay the footnote layout's fixpoint passes on the measured body in place instead of on copies of it. Every layout of a document with notes cloned the whole measured arena once, and once more per reservation pass; in wasm those copies cost most of the first edit after an open.
