---
'@betteroffice/agents': patch
'@betteroffice/docx': patch
---

Rendering a DOCX page to PNG in Node builds only that page, so documents of several thousand pages render on both engines instead of failing.
