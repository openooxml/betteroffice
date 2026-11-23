---
"@betteroffice/fonts": patch
---

Under Vite 7 and 8, an absent `@betteroffice/fonts-intl` or `@betteroffice/fonts-cjk` now counts as a missing add-on, as under esbuild and webpack: `isMissingFontAddon` recognises the error from `registerBundledFontFace` and `loadBundledFontBytes` instead of receiving Vite's raw resolver error.
