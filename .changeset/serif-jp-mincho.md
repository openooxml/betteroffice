---
'@betteroffice/fonts': patch
'@betteroffice/fonts-cjk': patch
---

Vendor Noto Serif JP and resolve Japanese Mincho requests to it instead of the sans face. MS Mincho, MS PMincho, Yu Mincho and the native ＭＳ 明朝 spellings now measure and render with the serif design they ask for, matching `fontResolver.ts`; Gothic, Meiryo, Meiryo UI and Yu Gothic requests still resolve to Noto Sans JP, and the `cjk-jp` script fallback still serves the sans face first.
