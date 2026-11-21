---
'@betteroffice/pptx': patch
'@betteroffice/pptx-native': patch
'@betteroffice/rust-crates': minor
---

Text drawn with substitute faces (Aptos, Verdana, Century Gothic) no longer gaps inside words. Canvas and native stretch each laid-out piece to its span. Display-list runs shaped without kerning carry `kerning: false`, and both engines draw them unkerned. Set `kerning` in `PositionedTextRun` literals.
