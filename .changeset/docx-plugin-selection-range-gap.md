---
"@betteroffice/docx-react": patch
---

Plugin selection display ranges across a frame replacement: an unchanged selection keeps its previous range until the new layout maps it, then one `selection-change` replaces or clears it. A moved selection stays empty until mapped. A superseded or failed read shows null, as before.
