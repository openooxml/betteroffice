---
"@betteroffice/docx": patch
---

Embedded images stay in the opened package and are read when a page first shows them, so image-heavy documents open faster and use less memory. `createCanvasImageResolver` takes an optional `media` source for the `media:{n}` image sources display lists now carry.
