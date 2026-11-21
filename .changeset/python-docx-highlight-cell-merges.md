---
'@betteroffice/python-docx': minor
---

`TextRun` gains `highlight`, and `TableCell` gains `grid_span` and `v_merge`: a run's highlight colour as written, the grid columns a cell spans, and `"restart"` or `"continue"` for a vertically merged cell. Each is `None` when the document does not set it.
