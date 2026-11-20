---
'@betteroffice/pptx': minor
'@betteroffice/pptx-react': patch
---

Breaking: `slideToPng` and its `SlideToPngOptions` are gone. The handle's `renderPng` replaces them and decodes images and paints with registered faces itself.

```ts
await slideToPng(handle.layoutSlide(i), { scale: 2, resolveImage }); // before
await handle.renderPng({ slide: i }, { scale: 2 }); // after
```
