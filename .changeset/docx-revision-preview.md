---
"@betteroffice/docx": minor
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": minor
---

Preview accept and reject decisions without changing the document. `YrsRenderEnv.revisionPreview` maps revision ids to `accepted` or `rejected`. A previewed insertion or deletion renders as plain text or is left out, at its original positions. Unlisted revisions keep their tracked-change markup. The preview is part of the render identity: a change relays out, the worker resyncs, and older frames are refused. Layouts and display-list queries record the preview they show. Undo history and revisions are untouched, and a paged export of a previewed layout is refused.
