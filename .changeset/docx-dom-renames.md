---
"@betteroffice/docx-react": patch
---

Renamed worker editor DOM hooks: `.paged-editor__peerless-input` is now `.paged-editor__input`, `data-peerless-docx` is `data-docx-worker-editor`, `data-peerless-pages-host` is `data-docx-pages-host`, and test id `peerless-paged-editor` is `docx-worker-paged-editor`. Update host CSS and tests that select them.
