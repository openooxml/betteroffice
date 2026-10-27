---
"@betteroffice/docx-react": patch
"@betteroffice/docx": patch
---

The caret moves with the text while you type: when a keystroke paints its letters ahead of layout, the caret is drawn right after them in the same frame, at the height of the real caret, and settles in one step if layout places it elsewhere. Moving or blinking the caret no longer causes a page layout.
