---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/xlsx-react': minor
'@betteroffice/agents': patch
'@betteroffice/rust-crates': minor
---

Conditional formats are decided once per edit, at most 10 ms a frame after rule compilation and range statistics, so frames can run over. Rule formula names parse once per frame. Display lists gain `conditionalPending`, cells left undecided; the editor asks for frames until 0.
