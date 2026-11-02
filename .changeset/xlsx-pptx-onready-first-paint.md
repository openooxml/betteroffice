---
'@betteroffice/xlsx-react': patch
'@betteroffice/pptx-react': patch
---

`XlsxEditor` and `PptxEditor` with `legacySyncApi` now call `onReady` once the first frame is painted. A hidden or zero-size editor calls it once it is shown.
