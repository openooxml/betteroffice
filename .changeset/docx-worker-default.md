---
"@betteroffice/docx-react": minor
---

The worker-owned editor is now the default; `experimentalWorkerOpen={false}` keeps the in-thread engine as a deprecated opt-in. Missing browser features, `mediaTokens`, `collaboration.initialUpdate`, and editable `document` sources without `documentBuffer` fall back with a console warning.
