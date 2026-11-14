---
'@betteroffice/docx-react': minor
'@betteroffice/xlsx-react': minor
'@betteroffice/pptx-react': minor
---

`resolveDocxShortcuts`, `resolveXlsxShortcuts` and `resolvePptxShortcuts` load from each package's `./shortcuts` subpath, keeping them out of the editor bundle. They were never released from the package root, so nothing breaks.
