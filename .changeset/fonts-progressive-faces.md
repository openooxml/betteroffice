---
"@betteroffice/fonts": minor
"@betteroffice/fonts-cjk": patch
---

Loaders of the sliced CJK faces gain an optional `progressive()` that opens the face from its skeleton and core and loads further outline chunks by glyph id, as a new `ProgressiveFontFace`; calling the loader still returns the whole face, and custom providers are unaffected.
