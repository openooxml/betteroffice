---
'@betteroffice/docx': minor
'@betteroffice/docx-native': minor
'@betteroffice/agents': minor
---

DOCX PNG export draws a placeholder for a picture it cannot draw instead of failing the page or leaving it out. The new `onDiagnostic` option reports each such picture and why. Agent page previews add a warning.
