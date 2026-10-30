---
'@betteroffice/pptx': patch
'@betteroffice/pptx-native': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-pptx': patch
---

Default fonts now load script coverage for unmapped families and mixed text when optional font packs are available. `requestedFontFacesJson` adds optional `scripts` per face, and presentations with only coverage fonts render successfully.
