---
"@betteroffice/docx-react": minor
---

Worker-only editing is now the default: the document lives in a worker; reads and edits are async. Deprecated `legacySyncApi` restores the sync hybrid and `onChange(Document)` until 0.6.0. `experimentalWorkerOpen={false}` and browsers lacking worker features use the in-thread engine.
