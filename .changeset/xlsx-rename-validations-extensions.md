---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
---

Renaming a sheet also rewrites data-validation, extension conditional-format and sparkline formulas that reference it, as Excel does. A rename those formulas cannot follow is refused before anything changes.
