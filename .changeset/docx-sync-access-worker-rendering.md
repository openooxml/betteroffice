---
"@betteroffice/docx-react": patch
---

With `experimentalWorkerOpen`, layout and rendering return to the worker after a synchronous editor call during load opens the document on the main thread, so later edits and tracked-change decisions stay fast. A one-time console warning names the synchronous call.
