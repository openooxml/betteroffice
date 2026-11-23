---
"@betteroffice/fonts": minor
"@betteroffice/docx": minor
"@betteroffice/pptx": minor
"@betteroffice/xlsx": minor
---

Browser pages load uninstalled font packs from jsDelivr; under a strict content policy, allow cdn.jsdelivr.net, install the packs, or call `configureFontCdn({ baseUrl: false })`. `@betteroffice/fonts-latin` is now an optional peer. WASM packages under Node use Base Latin unless packs are installed.
