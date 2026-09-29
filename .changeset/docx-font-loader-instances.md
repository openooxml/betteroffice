---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Keep font loading apart between editors on one page. A family that Google Fonts does not serve is requested once per page, and its failed stylesheet link is removed instead of accumulating on every mount. Embedded faces are keyed by their bytes: a second document that embeds different faces under the same family name registers them under an alias the editor paints with, and each editor releases its embedded faces and their blob URLs when it unmounts or loads its next document. `createFontLoadScope` and the editor's `onFontsLoaded` and `onError` now hear only their own instance's font loads, plus module-level ones.
