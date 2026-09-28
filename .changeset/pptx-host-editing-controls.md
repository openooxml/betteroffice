---
"@betteroffice/pptx": minor
"@betteroffice/pptx-react": minor
---

Expose host save interception, awaited input flushing, and pointer position queries. Add manual undo capture, explicit boundaries, and undoable comment repositioning. `PptxEditorApi.save()` now throws while input or a pointer gesture is pending (await `flushPendingInput()` first) and after the presentation is replaced, instead of returning bytes that miss accepted input.
