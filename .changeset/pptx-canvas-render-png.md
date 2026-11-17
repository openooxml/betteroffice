---
'@betteroffice/pptx': minor
'@betteroffice/pptx-react': patch
---

Breaking: `slideToPng` becomes the handle's `renderPng`, which decodes images and paints with registered faces itself.

```ts
await slideToPng(handle.layoutSlide(i), { scale: 2, resolveImage }); // before
await handle.renderPng({ slide: i }, { scale: 2 }); // after
```
