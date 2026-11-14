---
'@betteroffice/pptx': patch
'@betteroffice/pptx-native': patch
---

`renderPng` in the browser lays slides out with the bundled faces the native engine uses, so line breaks and row heights match native. On both engines, a needed face that cannot load now rejects with `PngExportError` (`font-unavailable`) naming the face.
