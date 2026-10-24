---
"@betteroffice/pptx-react": minor
---

**Breaking:** Worker-only editing is now the default, and `onReady` receives `PptxAsyncEditorApi`. Until 0.6.0, the deprecated `legacySyncApi` keeps the synchronous API on the in-thread engine; `experimentalWorkerOpen={false}` selects the deprecated in-thread engine with the async API. See https://docs.betteroffice.dev/docs/migrating-to-0-5.
