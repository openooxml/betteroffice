---
"@betteroffice/xlsx-react": patch
"@betteroffice/xlsx-i18n": patch
---

The experimental plugin edit gate now also covers paste and cut: plugins see each changed cell, pasted formulas already moved, with `source` `paste` or `cut`, and any refusal refuses the whole gesture before the clipboard or cells change.
