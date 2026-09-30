---
"@betteroffice/docx-react": patch
---

The layout worker now starts while the document's fonts are still loading, so the first worker layout no longer waits for it to boot afterwards.
