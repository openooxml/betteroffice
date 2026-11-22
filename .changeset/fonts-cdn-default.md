---
"@betteroffice/fonts": minor
"@betteroffice/docx": minor
"@betteroffice/pptx": minor
"@betteroffice/xlsx": minor
---

Browser pages now load uninstalled font packs from jsDelivr. With strict content policies, allow cdn.jsdelivr.net, install the packs, or call `configureFontCdn({ baseUrl: false })`. `@betteroffice/fonts-latin` is now an optional peer. WASM packages under Node lay out with Base Latin unless packs are installed.
