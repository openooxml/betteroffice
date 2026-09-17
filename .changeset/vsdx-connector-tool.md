---
"@betteroffice/vsdx": minor
---

Add `addConnectedShape` to the VSDX diagram handle: it inserts a shape and a
connector glued to an existing shape in one transaction and one undo entry, and
refuses the edit when the glue does not resolve to a connection point.
