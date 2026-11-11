---
"@betteroffice/docx-react": patch
---

Plugin selection display ranges that stay across a layout or frame replacement: an unchanged selection keeps its previous range until the new layout maps it, a moved one stays empty until mapped, and each emits a single `selection-change`.
