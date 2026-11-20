---
'@betteroffice/docx-react': patch
'@betteroffice/docx': minor
'@betteroffice/xlsx': minor
'@betteroffice/pptx': minor
---

A DOCX document opened with a seed and without collaboration now recovers after a worker restart that followed undo or redo; undo history starts at recovery. Adds public `clientId` and `onStream`, the `applyUpdate` command kind, optional engine `applyVerbatimUpdate`, and the binding's `apply_verbatim_update`.
