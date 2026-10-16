---
"@betteroffice/xlsx": patch
---

Applying a collaborator's update that arrives ahead of an earlier one no longer freezes the workbook. The update waits until the missing edit arrives, and every replica then shows the same cells.
