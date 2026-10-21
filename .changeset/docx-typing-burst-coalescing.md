---
"@betteroffice/docx-react": patch
---

Typing a burst of characters in a docx document sends the keys typed while one insertion is still being applied as a single insertion, so fast typing no longer queues one layout per key.
