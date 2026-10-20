---
"@betteroffice/xlsx-react": minor
---

**Breaking:** Worker-only editing is now the default, and `onReady` receives `XlsxAsyncEditorApi`. To keep the previous synchronous API until 0.6.0, pass the deprecated `legacySyncApi`; the deprecated `experimentalWorkerOpen={false}` still selects the in-thread engine.
