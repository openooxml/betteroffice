---
"@betteroffice/docx": patch
---

Canvas image rendering works in workers and bounds cached image memory by default. Adds optional `maxCacheBytes` to `createCanvasImageResolver` to customize the memory budget.
