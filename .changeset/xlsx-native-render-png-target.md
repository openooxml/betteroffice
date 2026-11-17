---
'@betteroffice/xlsx-native': minor
---

`renderPng` draws in the native engine through the new raw `renderPngTarget(targetJson, optionsJson)`, which honours scale and maxWidth and maxHeight for viewports and ranges. The raw `renderPng(viewportJson)` and `renderRangePng(args)` are removed.
