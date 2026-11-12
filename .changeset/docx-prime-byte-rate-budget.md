---
'@betteroffice/docx': patch
---

Background save seed priming now sizes each slice from the byte rate the last digesting slice measured. Slices that digest no bytes keep the budget, so the first story slice no longer starts at the 1 MiB cap.
