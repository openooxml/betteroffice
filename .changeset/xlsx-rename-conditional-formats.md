---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
---

Renaming a sheet rewrites the conditional-format rule formulas that reference it, as Excel does, instead of leaving them pointing at the old name. Rules kept as the file wrote them keep everything else they hold, undo restores the original text, and rules added in the session are no longer a reason to refuse the rename.
