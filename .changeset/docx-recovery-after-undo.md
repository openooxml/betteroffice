---
'@betteroffice/docx-react': patch
'@betteroffice/docx': minor
'@betteroffice/xlsx': minor
'@betteroffice/pptx': minor
---

A DOCX document opened with a seed and without collaboration now recovers after a worker restart that followed undo or redo, and undo history starts at recovery. Adds the public `clientId` and `onStream` members and the engine's `apply_verbatim_update`.
