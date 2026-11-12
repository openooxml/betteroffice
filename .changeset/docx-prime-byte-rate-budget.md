---
'@betteroffice/docx': minor
---

Background save seed priming now sizes each slice from the rate the last charging slice measured, growing at most a few times per slice. Slices that charge nothing keep the budget. Adds the `primeLastSliceCost` getter.
