---
"@betteroffice/docx-react": patch
---

With `experimentalWorkerOpen`, a layout pass that an edit made stale, or whose full layout the worker did not finish, runs again in the worker while it holds the document, instead of on the main thread.
