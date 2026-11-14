---
'@betteroffice/pptx': patch
'@betteroffice/pptx-native': patch
'@betteroffice/rust-crates': patch
---

Text drawn with a substitute face (Aptos, Verdana, Century Gothic) no longer gaps inside words: each piece the layout places is stretched to its span, alike on canvas and native. Display-list runs shaped without kerning carry `kerning: false`, and both engines draw them unkerned.
