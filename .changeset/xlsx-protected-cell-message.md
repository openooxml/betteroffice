---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-react': minor
'@betteroffice/xlsx-i18n': patch
---

Typing into a locked cell on a protected sheet now reports `sheet-protected` ("This cell is on a protected sheet.") instead of saying the entry could not be saved, and unlocked cells on protected sheets can be edited. `XlsxCommandFailureCode` and `XlsxCommandAdmissionCode` gain the member `sheet-protected`, so exhaustive switches over them need a case for it.
