---
'@betteroffice/xlsx-react': patch
'@betteroffice/pptx-react': patch
---

The in-thread `XlsxEditor` and `PptxEditor` now call `onReady` after the first frame is painted, as the worker editors already did, so `getPositionAtPoint` works inside it. A zero-size editor calls it once it gets a size.
