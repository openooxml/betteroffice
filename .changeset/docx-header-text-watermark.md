---
'@betteroffice/docx': minor
'@betteroffice/docx-native': patch
'@betteroffice/python-docx': patch
'@betteroffice/rust-crates': minor
---

Word text watermarks now draw on the pages whose header carries them, stretched to their shape box, and no longer on a page whose header has none. `TextWatermark` gains optional `widthEmu` and `heightEmu`.
