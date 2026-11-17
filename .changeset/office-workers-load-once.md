---
'@betteroffice/docx': patch
'@betteroffice/pptx': patch
'@betteroffice/xlsx': patch
---

In Safari and WebKit, apps that bundle the editors with ES module workers (Vite `worker.format: 'es'`) could run a worker twice and fail on its first command. Each worker entry now loads its code once.
