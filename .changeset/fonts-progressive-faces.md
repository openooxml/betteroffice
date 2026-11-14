---
"@betteroffice/fonts": minor
"@betteroffice/fonts-cjk": patch
---

Loaders of the sliced CJK faces gain an optional `slices` location, and the new `@betteroffice/fonts/progressive` entry opens it from the skeleton and core, loading outline chunks by glyph id; calling the loader still returns the whole face, and custom providers are unaffected.
