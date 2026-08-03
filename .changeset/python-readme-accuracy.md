---
"@betteroffice/python-xlsx": patch
"@betteroffice/python-pptx": patch
---

Both Python READMEs now describe what the bindings actually do. The
`betteroffice-xlsx` page said `save` regenerated the package and dropped parts
the model does not cover; saving has preserved charts, drawings, pivot tables,
comments, macros and custom XML since part preservation landed, so the Status
section now states the preservation and the limits that remain — an edited sheet
is still reserialized from the model, and this binding exposes no structural
edits at all. The `betteroffice-pptx` page said unregistered families fell back
to a metrics-only path; `render_slide` raises `RenderError: no font has been
registered for slide text` instead, so the layout section leads with that error
and then shows registering a face, and says what a family you did not register
resolves to once one exists. That page also told readers to `pip install
betteroffice-pptx`, which is not on PyPI; it now gives the editable install from
a checkout.

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
