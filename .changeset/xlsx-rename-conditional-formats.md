---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
---

Renaming a sheet rewrites the conditional-format formulas that reference it, as Excel does. Rules from the file keep everything else they hold, undo restores them, and rules added in the session no longer block the rename.
