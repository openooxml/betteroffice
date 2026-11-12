---
'@betteroffice/docx': minor
---

Background save seed priming now sizes each slice from the byte rate the last digesting slice measured. Slices that digest no bytes keep the budget. Adds the `primeLastSliceBytes` getter.
