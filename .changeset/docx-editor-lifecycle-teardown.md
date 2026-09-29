---
"@betteroffice/docx-react": patch
---

Editor teardown is more precise: a StrictMode remount keeps its worker, pages whose canvases no worker took repaint, a document that fails to open frees its session at once, and decoded images are released with their document.
