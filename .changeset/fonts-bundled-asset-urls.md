---
'@betteroffice/fonts': patch
---

fonts: the bundled entry now loads every face the package ships, so Inter, Roboto, Gelasio, Comic Relief and the other newly added families no longer reject with `Unknown bundled font asset`, and bundlers emit their files. A missing non-CJK face is reported as an unknown asset instead of asking for `@betteroffice/fonts-cjk`.
