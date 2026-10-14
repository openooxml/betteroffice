---
"@betteroffice/docx-react": patch
---

Worker-only DOCX pages that cannot be painted no longer make the editor re-request full frames in a loop. It retries once, waits for its own pending request, reports the stall once through `onError` without closing the document, and retries after zoom or viewport changes.
