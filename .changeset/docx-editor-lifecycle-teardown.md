---
"@betteroffice/docx-react": patch
---

Tear the editor down more precisely. A StrictMode remount no longer lets the destroyed first worker's failure take down the worker the second mount started, and page canvases that were handed over for worker painting but that no worker took remount on the main-thread path instead of staying blank. An editing session whose document fails to open is freed at once, and each document's decoded images are dropped when the editor moves to the next one.
