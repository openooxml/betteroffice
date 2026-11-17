---
"@betteroffice/docx": minor
"@betteroffice/docx-native": patch
"@betteroffice/rust-crates": minor
---

DOCX PNG export names a picture whose part is missing by its relationship id, in `onDiagnostic` and in the `strictPictures` rejection, instead of leaving it unnamed. Display-list pictures carry it as `relationshipId`.
