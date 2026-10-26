---
"@betteroffice/pptx": patch
"@betteroffice/pptx-react": patch
"@betteroffice/desktop": patch
---

Presentations open and show every slide when font downloads fail. Text falls back to available fonts and the editor reports the fallback once through `onError` as a recoverable error.
