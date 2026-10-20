---
"@betteroffice/pptx-react": minor
---

**Breaking:** Worker-only editing is now the default. To keep the previous synchronous API until 0.6.0, pass the deprecated `legacySyncApi`; the deprecated `experimentalWorkerOpen={false}` still selects the in-thread engine.
