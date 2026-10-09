---
"@betteroffice/docx-react": minor
---

Worker-open editors load their editing copy eagerly and replay input typed while it loads instead of dropping it; until it is ready, synchronous ref reads return empty answers and synchronous edits throw `DocxReplicaNotReadyError`.
