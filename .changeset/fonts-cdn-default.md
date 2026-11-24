---
"@betteroffice/fonts": minor
"@betteroffice/docx": minor
"@betteroffice/pptx": minor
"@betteroffice/xlsx": minor
---

Browser pages load font packs that are not installed (`@betteroffice/fonts-intl`, `@betteroffice/fonts-cjk`) from jsDelivr; under a strict content policy, allow cdn.jsdelivr.net, install the packs, or call `configureFontCdn({ baseUrl: false })`. A blocked CDN substitutes as before, with one console warning per pack.
