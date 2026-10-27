---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
---

Adds `headerFooter` facts to plugin `readCells` sheet entries: printable left, center and right sections and the active first, odd and even variants; null when a sheet has none. Workbook snapshots written by this version do not decode in 0.5.1.
