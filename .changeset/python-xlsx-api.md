---
"@betteroffice/python-bindings": patch
---

The Python binding now reaches collaboration, history, agent proposals, formatting and sheet metadata instead of stopping at open, read and save. `open_collaborative`, `state_vector`, `state_as_update`, `diff` and `apply_update` expose the Yrs primitives directly, so replicas converge over any transport without the package taking on an event loop. `undo`, `redo`, `history` and `set_many` give a local undo stack where a batch is one step and a peer's update stays out of it. `propose`, `proposals`, `accept_proposal` and `reject_proposal` carry the before and after text of every proposed edit and write nothing to the sheet until acceptance. `set_style` and `set_number_format` apply across a range and refuse an unknown alignment rather than ignoring it. `active_sheet`, `set_active_sheet`, `merged_ranges` and `last_calculation` read metadata the engine already tracked.

Breaking. `set` returns `Mutation` instead of `bool`, matching every other mutating call. `Mutation.__bool__` is `applied`, so `if wb.set(...)` reads the same; `wb.set(...) is True` no longer holds. `Mutation.changed` lists the cells the engine recalculated, which does not include a cell written directly.
