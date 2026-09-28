---
'@betteroffice/vsdx': patch
---

Add `setShapeData`, a single engine op that writes a batch of shape-data row values for one shape as one undo entry. Every row is decided before anything is written, so a refusal anywhere leaves the document untouched, and the caller gets one receipt per row saying whether it was written and why not. Date, duration and currency rows keep their typed value rather than being rewritten as text, a number row takes a numeric value, and each write carries a non-empty formula.
