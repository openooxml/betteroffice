---
'@betteroffice/vsdx': patch
---

Add `setShapeData`, a single engine op that writes a batch of shape-data row values for one shape as one undo entry. Every row is decided before anything is written, so a refusal anywhere leaves the document untouched, and the caller gets one receipt per row saying whether it was written and why not. Date, duration and currency rows refuse rather than being rewritten as text, a number row refuses a non-numeric value, and an empty formula is rejected.
