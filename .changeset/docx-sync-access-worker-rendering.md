---
"@betteroffice/docx-react": patch
---

With `experimentalWorkerOpen`, layout stays in the worker after a synchronous ref call during load (a one-time warning names it), and a worker that fails after load is replaced; a repeated failure reaches `onError`.
