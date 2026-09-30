---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": minor
"@betteroffice/docx-react": patch
---

Adds opt-in `set_windowed_incremental_builds` to limit incremental rebuilds to the display window and caret pages. The editor enables this for edits and proposal decisions, rebuilds other affected pages in the background, and waits for exact visible proposal geometry.
