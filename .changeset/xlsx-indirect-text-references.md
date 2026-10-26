---
"@betteroffice/xlsx": patch
---

INDIRECT now resolves references written in lowercase and in R1C1 style (`INDIRECT("R5C2", FALSE)`), and a formula pointing at a deleted cell on another sheet (`'Sheet'!#REF!`) shows #REF! as in Excel, instead of a blank or a stale value.
