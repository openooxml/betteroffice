---
"@betteroffice/vsdx": minor
---

Add `addFreeConnector` to the VSDX diagram handle: it inserts a connector glued
at its begin only, leaving the end where the draft puts it, and refuses the edit
when the glue does not resolve to a connection point.
