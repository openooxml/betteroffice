---
"@betteroffice/docx-react": patch
---

Replacing or closing a document no longer throws "null pointer passed to rust" when the previous document's pages rebuild, or its fonts finish loading, after the switch.
