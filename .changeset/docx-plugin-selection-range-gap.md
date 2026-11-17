---
"@betteroffice/docx-react": patch
---

In the in-thread editor, a plugin's selection display range no longer drops to null when the layout is replaced under an unchanged selection; it keeps its previous range until the new layout maps it, then one `selection-change` follows.
