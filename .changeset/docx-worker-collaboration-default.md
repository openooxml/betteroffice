---
"@betteroffice/docx": minor
"@betteroffice/docx-react": minor
---

`collaboration` now uses the default worker editor, with `onChange` refused as elsewhere. `onReplica` receives an async replica: read through `replica.async`, or pass `legacySyncApi` for synchronous reads. Unpinned client ids are fresh per open; pinned ids from `Number.MAX_SAFE_INTEGER - 1` are rejected.
