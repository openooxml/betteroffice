# @betteroffice/python-xlsx

## 0.0.2

### Patch Changes

- 98b9225: Both distributions point their PyPI `Documentation` link at
  https://docs.betteroffice.dev/docs/python, a page that names the package, gives
  the install line and a first example. The link used to land on the documentation
  root, which never mentioned Python at all.
- bdedb87: Both Python READMEs now describe what the bindings actually do. The
  `betteroffice-xlsx` page said `save` regenerated the package and dropped parts
  the model does not cover; saving has preserved charts, drawings, pivot tables,
  comments, macros and custom XML since part preservation landed, so the Status
  section now states the preservation and the limits that remain — an edited sheet
  is still reserialized from the model, and this binding exposes no structural
  edits at all. The `betteroffice-pptx` page said unregistered families fell back
  to a metrics-only path; `render_slide` raises `RenderError: no font has been
registered for slide text` instead, so the layout section leads with that error
  and then shows registering a face, and says what a family you did not register
  resolves to once one exists.

  Both READMEs name the import next to the install line, because `pip install
betteroffice-xlsx` gives `import betteroffice_xlsx`, and both name the incumbent
  they are compared against in the opening paragraph rather than a hundred lines
  down. The xlsx API table stops presenting the static `Workbook.open_collaborative`
  as an instance method, and gains `value`, `formula`, `proposals`,
  `merged_ranges`, `last_calculation`, `sheet_index`, `can_undo`/`can_redo`, and
  the `now_serial` keyword that supplies the clock `TODAY()` and `NOW()` read.
  `StaleProposalError` now documents its way out, `accept_proposal(id,
force=True)`. The proposals example prints the value it actually produces, and
  the pptx snippets no longer assume the first slide has a shape or that the shape
  bears text.

  Both distributions declare `Operating System :: OS Independent` and per-minor
  `Programming Language :: Python` classifiers for 3.9 through 3.13, so PyPI's
  version filter finds them, swap the `openpyxl-alternative` and
  `python-pptx-alternative` keywords nobody searches for the bare project names,
  and add a `Changelog` project URL.

- ab39d50: The Python binding now reaches collaboration, history, agent proposals, formatting and sheet metadata instead of stopping at open, read and save. `open_collaborative`, `state_vector`, `state_as_update`, `diff` and `apply_update` expose the Yrs primitives directly, so replicas converge over any transport without the package taking on an event loop. `undo`, `redo`, `history` and `set_many` give a local undo stack where a batch is one step and a peer's update stays out of it. `propose`, `proposals`, `accept_proposal` and `reject_proposal` carry the before and after text of every proposed edit and write nothing to the sheet until acceptance. `set_style` and `set_number_format` apply across a range and refuse an unknown alignment rather than ignoring it. `active_sheet`, `set_active_sheet`, `merged_ranges` and `last_calculation` read metadata the engine already tracked.

  Breaking. `set` returns `Mutation` instead of `bool`, matching every other mutating call. `Mutation.__bool__` is `applied`, so `if wb.set(...)` reads the same; `wb.set(...) is True` no longer holds. `Mutation.changed` lists the cells the engine recalculated, which does not include a cell written directly.
