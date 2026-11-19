---
'@betteroffice/docx-react': patch
'@betteroffice/docx': minor
'@betteroffice/xlsx': minor
'@betteroffice/pptx': minor
---

Recovery after a worker restart no longer fails, or brings back text that was undone, when undo or redo was used after a checkpoint. After a recovery, the undo history starts at the recovery.
