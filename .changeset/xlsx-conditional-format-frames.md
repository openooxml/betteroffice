---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/xlsx-react': minor
'@betteroffice/agents': patch
'@betteroffice/rust-crates': minor
---

Conditional formats evaluate once per edit within a frame budget; after a direct edit a cell keeps its previous format until decided. Rule formula names parse once per frame. Display lists gain `conditionalPending`, cells left undecided; the editor asks for frames until 0.
