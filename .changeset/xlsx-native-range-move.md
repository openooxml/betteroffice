---
'@betteroffice/xlsx': minor
'@betteroffice/rust-crates': minor
---

Adds `moveRange`, also on the worker session, which moves a range within one sheet and rewrites formulas that refer to it in one undo step. Partial ranges and features a move cannot preserve are refused.
