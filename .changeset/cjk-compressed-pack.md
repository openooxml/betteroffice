---
"@betteroffice/fonts-cjk": minor
"@betteroffice/fonts": minor
---

The optional CJK add-on now ships compressed `.otf.zst` assets, decoded automatically by the font manager. `CJK_FONT_ASSET_URLS` keys and `BundledFontFace.file` use the packed filenames; custom asset hosts must serve them.
