---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
---

The in-thread XLSX editor now calls `onChange` when a collaborator's edit changes the workbook, including a full-state update into a freshly opened replica, like the worker editor. `WorkbookHandle.onUpdate` listeners also receive `{ fromPeer }`.
