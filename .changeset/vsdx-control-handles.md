---
"@betteroffice/vsdx": patch
"@betteroffice/vsdx-react": patch
---

Resolve the `Control` section and draw its yellow diamond handles on the selected shape, positioned through the shape's own transform and its group ancestors. Dragging a handle writes the row's `X` and `Y` through the mutation policy in a single transaction, so the gesture is one undo entry and a guarded axis refuses the whole drag instead of committing half of it. Handles whose `XCon`/`YCon` behaviour is hidden are not drawn, and an axis-locked handle keeps that axis fixed. Core adds `setControlHandle`.
