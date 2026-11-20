---
'@betteroffice/docx-react': patch
'@betteroffice/docx': minor
'@betteroffice/xlsx': minor
'@betteroffice/pptx': minor
---

A DOCX document opened with a seed and without collaboration now recovers after a worker restart following undo or redo; undo history starts at recovery. Adds `clientId`, `onStream`, the `applyUpdate` command kind, and optional stream members: the host's `recordUpdate`, `takeStream`, `withCommandWork`, and the engine's `applyVerbatimUpdate`.
