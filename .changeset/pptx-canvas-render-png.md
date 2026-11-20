---
'@betteroffice/pptx': minor
'@betteroffice/pptx-react': patch
---

Breaking: `slideToPng` and `SlideToPngOptions` are gone. The handle's `renderPng` replaces them: it resolves to a `Uint8Array`, not a `Blob`, and decodes images itself.

```ts
await slideToPng(handle.layoutSlide(i), { scale: 2, resolveImage }); // before
await handle.renderPng({ slide: i }, { scale: 2 }); // after
```
