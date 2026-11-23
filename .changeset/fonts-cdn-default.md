---
"@betteroffice/fonts": minor
"@betteroffice/docx": minor
"@betteroffice/pptx": minor
"@betteroffice/xlsx": minor
---

Browsers fetch missing packs from jsDelivr by default. `@betteroffice/fonts-latin` is optional.

Breaking: `loadBundledFontBytes` rejects instead of returning a substitute when a face is unavailable after a failed default CDN fetch or because `@betteroffice/fonts-intl` is absent.
