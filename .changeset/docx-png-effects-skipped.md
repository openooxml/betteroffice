---
'@betteroffice/docx': minor
'@betteroffice/docx-native': minor
'@betteroffice/xlsx': minor
'@betteroffice/pptx': minor
'@betteroffice/agents': patch
---

DOCX PNG export tells onDiagnostic, once per page, which kinds of effect the engine does not draw (`effects-skipped`). The shared PNG diagnostic type gains that code, so a host with an exhaustive switch on the code gets a new case.
