# @betteroffice/docx

## 0.0.3

### Patch Changes

- b34bb01: Docx typing hot path is 7x faster (resident region fast path, memoized font parsing, direct frame-delta encoding, incremental worker sync); pages no longer remount and flash on remote or structural edits; page bitmaps are windowed to the viewport on long documents; the caret is painted by the renderer in the same frame as the glyphs while typing and blinks in the DOM at idle.

## 0.0.2
