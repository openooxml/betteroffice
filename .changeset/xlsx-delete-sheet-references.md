---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/rust-crates': minor
---

Deleting a sheet turns references to it into `#REF!` in formulas, defined names, conditional formats and data validations, as Excel saves them; 3-D spans narrow. Chart spans no longer change when a sheet inside them goes. Undo restores the text.
