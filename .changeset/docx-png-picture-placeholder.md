---
'@betteroffice/docx': minor
'@betteroffice/docx-native': minor
'@betteroffice/agents': minor
'@betteroffice/rust-crates': minor
---

DOCX PNG export draws a placeholder for a picture it cannot draw instead of rejecting the page or leaving it blank; `onDiagnostic` reports each one and why. `strictPictures: true` keeps rejecting. Agent previews warn.
