---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Documents opened with `experimentalWorkerOpen` save in the document worker, so viewing sessions save and download without loading the document on the main thread and editors no longer pause while saving.
