---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/xlsx-react': minor
'@betteroffice/agents': patch
'@betteroffice/rust-crates': minor
---

Conditional formats evaluate once per edit within a frame budget. Display lists gain `conditionalPending`, the visible cells a frame left undecided; the editor asks for another frame until it is 0, and exports always decide every cell.
