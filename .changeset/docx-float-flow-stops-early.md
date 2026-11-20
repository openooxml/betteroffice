---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Typing in a document with wrapped pictures, shapes or floating tables re-measured every block up to the next page or section break. It now re-measures from the edited block and stops once the text flow beside the floats is back where it was, with the same result as a fresh layout.
