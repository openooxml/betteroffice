---
"@betteroffice/xlsx": patch
---

Structured references to a table column whose header starts or ends with a space, such as `Table1[[#This Row],[ Total]]`, resolve to that column instead of `#REF!`, and keep the space when the formula is rewritten.
