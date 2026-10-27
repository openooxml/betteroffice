---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
---

Adds `headerFooter` facts to plugin `readCells` sheet entries by default, including printable sections and active page variants. Known absent headers and footers are null; unknown source facts are omitted.
