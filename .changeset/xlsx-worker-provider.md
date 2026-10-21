---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
---

Worker editors now run `CollaborationProvider` in the workbook worker by default, batching incoming transport frames into one session command. Existing collaboration and awareness APIs remain compatible.
