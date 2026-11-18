---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/xlsx-react': minor
'@betteroffice/agents': patch
'@betteroffice/rust-crates': minor
---

Conditional formats evaluate once per edit within a frame budget, showing their previous result until decided; names in rule formulas parse once per frame. Display lists gain `conditionalPending`, visible cells a frame left undecided; the editor asks for frames until 0; exports decide every cell.
