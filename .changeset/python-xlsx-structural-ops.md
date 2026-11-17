---
'@betteroffice/python-xlsx': minor
---

Adds `apply_ops` for the raw structural ops TypeScript's `applyOps` takes, typed in `betteroffice_xlsx.ops`, and `move_range`, `set_hyperlink`, `remove_hyperlink`, `edit_charts`, `move_chart`, `capture_format` and `apply_format`. Refusals raise `OperationError` or `CollaborativeStructureError` with TypeScript's message.
