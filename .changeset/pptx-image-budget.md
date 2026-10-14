---
"@betteroffice/pptx": patch
"@betteroffice/pptx-react": patch
---

Pictures use bounded decode sizes and image caching by default, reducing memory use in picture-heavy presentations. Adds optional `maxDimension` to `decodePresentationImage` and `acquire` and `release` to `CanvasImageResolver` for scoped painting.
