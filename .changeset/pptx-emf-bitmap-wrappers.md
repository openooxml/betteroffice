---
'@betteroffice/pptx': patch
---

Render EMF pictures that wrap a bitmap: `presentationImageBlob` unwraps an enhanced metafile whose only ink is one unscaled `EMR_STRETCHDIBITS` covering its bounds into the BMP it carries, alongside the bitmap-only WMF wrappers already handled. The blit must be an unrotated, uncropped `SRCCOPY` of a `BI_RGB` DIB that fills the metafile frame; metafiles carrying vector ink, a second blit or a scaled source stay untouched.
