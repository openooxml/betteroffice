---
'@betteroffice/pptx': patch
'@betteroffice/pptx-native': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-pptx': patch
---

Default fonts now load CJK script coverage for unmapped families and mixed text when the CJK packs are available; other scripts keep the browser font. `requestedFontFacesJson` adds optional `scripts` per face, and presentations with only coverage fonts render successfully.
