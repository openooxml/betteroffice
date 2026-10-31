---
'@betteroffice/pptx-native': patch
---

`@betteroffice/pptx-native` now installs `@betteroffice/fonts-latin`, `@betteroffice/fonts-intl` and `@betteroffice/fonts-cjk` and reads their faces from disk on demand, so Latin, international and CJK text renders with real faces without any fetch. On by default.
