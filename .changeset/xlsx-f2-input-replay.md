---
"@betteroffice/xlsx-react": patch
---

Fixes F2 cell editing in the default worker editor so keys pressed while the cell input loads apply to its existing value and preserve the caret position. Saving or enabling `readOnly` before loading completes leaves the cell unchanged.
