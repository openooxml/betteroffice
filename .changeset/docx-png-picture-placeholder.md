---
'@betteroffice/docx': minor
'@betteroffice/docx-native': minor
'@betteroffice/agents': minor
---

DOCX PNG export draws a placeholder for a picture it cannot draw instead of rejecting the page; `onDiagnostic` reports each one and why. Pass `strictPictures: true` to keep rejecting. Agent page previews add a warning.
