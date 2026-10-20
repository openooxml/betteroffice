---
"@betteroffice/docx-react": minor
---

**Breaking:** Worker-only editing is now the default, and the editor ref's document reads and edits are async. To keep the previous synchronous API until 0.6.0, pass the deprecated `legacySyncApi`; the deprecated `experimentalWorkerOpen={false}` still selects the in-thread engine.
