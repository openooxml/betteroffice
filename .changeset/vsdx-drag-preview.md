---
"@betteroffice/vsdx": patch
"@betteroffice/vsdx-react": patch
---

Track the pointer while dragging or Shift-drag resizing a shape, painting a live outline of the
landing position on the editor's overlay canvas instead of moving the shape only on release.

Add `modelPointToCanvas`, the exact forward of `canvasPointToModel`, and commit a gesture only once
it passes the same drag threshold the preview uses, so the preview and the committed geometry can
never disagree.
