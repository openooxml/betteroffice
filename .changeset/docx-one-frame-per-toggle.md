---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
"@betteroffice/docx-react": patch
---

With `set_windowed_incremental_builds` on, a full display rebuild also builds only the display window and the caret's page. The editor now waits for a pending frame before building more pages, so a proposal decision builds one frame.
