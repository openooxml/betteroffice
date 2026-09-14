---
"@betteroffice/vsdx-react": patch
"@betteroffice/vsdx-i18n": patch
---

Make the editor canvas focusable and give it a keyboard layer: undo and redo, Delete or Backspace,
arrow-key nudge with a larger Shift step, and Escape to cancel a gesture or clear the selection.
Keys are ignored while focus is in a text input, and a handle resize of a locked shape is refused
with a translated message.
