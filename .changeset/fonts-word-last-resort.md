---
'@betteroffice/fonts': minor
'@betteroffice/docx': patch
---

`resolveLastResortFace` and the provider's `resolveLastResort` take an optional `office`: `'powerpoint'`, the default, substitutes Calibri for an unknown sans family and `'word'` substitutes Arial. DOCX layout follows Word.
