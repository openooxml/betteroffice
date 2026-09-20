---
'@betteroffice/pptx': patch
---

Render slides that carry a picture the browser cannot decode: `paintSlide` treats a rejected `resolveImage` as an unresolved picture, so the remaining primitives on that slide, and every later slide, still paint. A resolver that rejects — `createImageBitmap` raises `InvalidStateError: The source image could not be decoded` for EMF, vector WMF and other media no browser decodes — previously rejected the whole `paintSlide` call, blanking the slide. The picture's own outline still strokes, matching an unresolved asset.
