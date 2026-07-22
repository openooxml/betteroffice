---
"@betteroffice/rust-crates": patch
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
---

Render pending proposals as Word-style tracked changes: struck-through old
values with a red run highlight, new values in green with a dashed underline
and green run highlight, laid out side by side or new-over-old and following
cell alignment. Proposal staging recalculates the formula graph and ghosts
downstream dependents whose computed values change, proposal edits can carry
a number format, and no-op proposals render unmarked.
