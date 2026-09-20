---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
---

Recognise `TRANSPOSE` in the formula engine. A 1x1 argument — a single cell, a literal, or any scalar expression — transposes to itself, with a blank cell transposing to `0`; errors propagate. The evaluator has no array value type, so a multi-cell argument returns `#VALUE!` rather than a silently wrong scalar, the same answer a bare range already gets in scalar context. Full `m x n` transposition waits on an array result representation.
