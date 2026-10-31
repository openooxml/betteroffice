---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-react': minor
'@betteroffice/xlsx-i18n': patch
---

Typing into a locked cell on a protected sheet now reports `sheet-protected` ("This cell is on a protected sheet.") instead of a failed save; unlocked cells there can be edited. `XlsxCommandFailureCode` and `XlsxCommandAdmissionCode` gain `sheet-protected`, so exhaustive switches need a case.
