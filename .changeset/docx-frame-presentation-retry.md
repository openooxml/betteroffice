---
"@betteroffice/docx-react": patch
---

Worker-only DOCX pages that cannot be painted, for example when the browser refuses a canvas 2D context, no longer make the editor re-request full frames in a loop. The editor requests a full frame and retries once, counting only failures of the frames it requested, then reports the failure once through `onError` and keeps the document open. It tries again after a zoom, device pixel ratio or viewport change, and recovers when a later frame installs.
