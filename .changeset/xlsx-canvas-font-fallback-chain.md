---
'@betteroffice/xlsx': patch
---

Canvas PNG exports without `fonts` now draw text in Korean, Japanese, Chinese, Arabic and other scripts through the same bundled fallback faces as the native engine, in its order, from installed font packages only. Sheets with Latin text only load nothing more.
