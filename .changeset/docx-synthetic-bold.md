---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

DOCX bold requests now synthesize weight when the selected font face has no intrinsic bold face, while preserving glyph advances.
