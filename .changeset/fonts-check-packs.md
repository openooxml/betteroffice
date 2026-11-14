---
'@betteroffice/fonts': patch
'@betteroffice/rust-crates': patch
---

`checkFontPacks()` and `BundledFonts::pack_diagnostics()` report installed font packs that cannot be used, with the pack and the reason, instead of skipping them silently.
