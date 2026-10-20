---
"@betteroffice/docx": minor
"@betteroffice/docx-react": minor
---

`onReplica` receives an async replica: read through `replica.async`, or pass `legacySyncApi` for synchronous reads. Unpinned client ids are fresh per open, a pinned id that would reuse clocks opens fresh with a warning, and ids from `Number.MAX_SAFE_INTEGER - 1` are rejected.
