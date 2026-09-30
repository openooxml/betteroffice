---
"@betteroffice/docx-react": patch
---

Replacing a document no longer throws "null pointer passed to rust" when the previous document's pages rebuild before the new one's first layout.
