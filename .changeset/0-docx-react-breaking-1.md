---
"@betteroffice/docx-react": minor
---

**Breaking:** Worker-only editing is now the default, and the editor ref's document reads and edits are async. Until 0.6.0, the deprecated `legacySyncApi` keeps the synchronous ref on the hybrid engine, and the deprecated `experimentalWorkerOpen={false}` keeps the 0.4 in-thread engine. See https://docs.betteroffice.dev/docs/migrating-to-0-5.
