---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Docx typing hot path is 7x faster (resident region fast path, memoized font parsing, direct frame-delta encoding, incremental worker sync); pages no longer remount and flash on remote or structural edits; page bitmaps are windowed to the viewport on long documents; the caret is painted by the renderer in the same frame as the glyphs while typing and blinks in the DOM at idle.
