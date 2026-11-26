---
"@betteroffice/pptx-react": patch
---

`onError` now reports refused input as the exported `PptxCommandRefusedError` (`code: 'command-refused'`) instead of `PptxWorkerEditorFailedError`. A click queued behind refused notes or comment edits is no longer refused.
