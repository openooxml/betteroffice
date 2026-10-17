---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Tables that two collaborators insert at the same time keep their own cells, so undoing or deleting one no longer breaks the other. Cell story ids are opaque: a cell created by an edit may carry a `~<n>` suffix, for example `body:t0:r0c0~51`.
