---
"@betteroffice/docx-react": patch
---

Worker-only DOCX pages that cannot be painted, for example when the browser refuses a canvas 2D context, no longer make the editor re-request full frames in a loop. The editor retries once, reports the failure once through `onError`, and tries again after a zoom or viewport change or when the worker sends a frame that installs.
