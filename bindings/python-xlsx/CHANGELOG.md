# @betteroffice/python-xlsx

## 0.2.0

### Minor Changes

- 67e0ae4: Release the accumulated DOCX, PPTX, and XLSX engine improvements in the Python bindings as minor updates.

### Patch Changes

- e7f4868: Honor Text cell formatting when entering values, including numeric strings, booleans, and formula-looking text. General cells retain Excel-style input coercion. Apply the same format-aware behavior to cell assignments, batches, and proposals.
- d1c6a55: xlsx: recalculation now supports `SUBTOTAL`, `AGGREGATE` and `XMATCH`. `SUBTOTAL` selects one of eleven aggregates by code, with the 100-series treated the same as the 1-series since manually hidden rows are not modelled. `AGGREGATE` adds `LARGE`, `SMALL`, `PERCENTILE` and `QUARTILE` on top, and its error-ignoring options recompute over the cells that are not errors. `XMATCH` supports exact, next-smaller and next-larger match modes and a reverse search.
- 3153c95: xlsx: row heights stored against a cached `sheetFormatPr/@defaultRowHeight` now render at the height Excel recomputes for the sheet's normal font
- 6bd3783: xlsx: recalculation now reads an argument that computes an array wherever a range is accepted, resolves whole-row and reference-range references, and answers per element where Excel does. Together with twenty-three functions that had no implementation — `DATEVALUE`, `FV`, `TREND`, `FREQUENCY`, `YEARFRAC`, `NORM.DIST`, `ADDRESS`, `HYPERLINK`, the `.INTL` workday pair and the `-A` aggregates among them — this covers the idioms real workbooks are built from: `MEDIAN(IF(range=key,range))` and the other CSE aggregates, `INDEX(data,MATCH(key,a&b,0))` over concatenated keys, `AGGREGATE(15,6,ROW(range)/(range=key),n)` for the nth match, `SUMPRODUCT(--(range=key))`, `ROWS($1:1)` as a running counter, and `AA9:INDEX(AA9:AH9,n)` as a computed endpoint.

  `INDEX` now picks its reference or array form from its first argument as Excel does, `MATCH` answers once per key, and a formula whose result is an empty reference stores 0 rather than staying blank. A whole-column reference costs the rows the sheet actually reaches, so a handful of `COUNTIFS(F:F,…)` formulas no longer exhaust the recalculation budget and leave the rest of the workbook reporting `#NUM!`.

  On the 165-workbook calculation benchmark, which strips cached values and compares against Excel's own full rebuild, this moves the corpus from 92.0076% to 99.5993% of cells matching Excel — 133 of 165 workbooks now match exactly, up from 49 — with no workbook regressing. LibreOffice scores 85.8713% on the same corpus.

- 21d83b3: Keep a column style attached to its columns through an edit, and make it part of what identifies a workbook. Inserting or deleting columns now moves the `<col>` runs with the cells: a run the insert splits widens over the new columns, a run to its right shifts, and a delete clips a run to the columns that survive and drops one it consumes whole. The collaboration fingerprint hashes column styles at the current schema version, so two peers whose workbooks differ only there no longer share a session and render differently; a workbook that declares no column style keeps the fingerprint every earlier release computed, and one that declares some carries that pre-change fingerprint in its accepted set, so saved snapshots still attach. A worksheet may declare at most 65,536 style-bearing `<col>` runs, refused as `TooManyColumnStyles` in the shape of the existing cell, hyperlink and chart caps.
- 21d83b3: Give a cell the format its column declares. A `<col>` run's `style` is now parsed, including on a run that sets no width, which every declaration used to be dropped from, and a cell with no `s` of its own takes it. A column style that names a solid fill now paints the whole column, so a sheet tinted through its columns arrives tinted instead of white, and a column filled white covers the gridlines Excel covers there. A cell's own style still wins, and an unsized row fits its content through the column style when the cell names none. On the eight XLSX fidelity samples this moves `bootstrapping-calculator` from 0.72960 to 0.90937 and `wazobia-valuation-model` from 0.88167 to 0.90700, with the other six unchanged and every page range still exact. Column styles stay a render projection: nothing is written back, so an untouched sheet still re-serializes byte-identically.
- 91554bf: Restore inherited column formatting on undo and redo of structural edits. History retains the original ordered column-style runs, including overlapping, clipped and entirely deleted runs, so restored columns and sheets render with their original formatting.
- d1c6a55: xlsx: recalculation now supports `LINEST` over one or more predictors, with the full statistics block its fourth argument asks for, and a scalar cell reading an array-only result through `INDEX` now evaluates it as an array instead of returning `#VALUE!`. The fit uses Householder QR rather than the normal equations, so a polynomial design such as `LINEST(y, x^{1,2,3,4,5,6})` stays conditioned. On the 165-workbook calculation benchmark this takes `sheetpedia-def4407b7c16` from 0 to 1989 of 1989 cells and `sheetpedia-c70e22c95a64` from 36 to 93 of 97, moving the corpus from 88.4700% to 90.1053% of cells matching Excel with no workbook regressing.
- b4db492: xlsx: a row with no explicit height now fits the workbook's normal font. A cell that names no style used to be measured at a fixed 11pt while the text in it was drawn at the normal size, so a workbook whose normal font is Arial 10 or Calibri 12 got a row grid that disagreed with its own text. When printing, the row fit and the text now read the same resolved face, so the two cannot diverge.
- d1c6a55: xlsx: `RANK`, `RANK.EQ`, `RANK.AVG` and `MATCH` given a range of values now answer one result per value instead of a single error, which is the shape the classic `LOOKUP(2, 1/(RANK(col,col,1)=n), col)` sorting idiom depends on. Their reference argument stays a range. On the 165-workbook calculation benchmark this takes `sheetpedia-eb5de1652388` from 133 to 945 of 946 cells, moving the corpus from 90.1053% to 90.7508% with no workbook regressing.
- d1c6a55: xlsx: recalculation now supports `QUOTIENT`, `SERIESSUM`, `ISEVEN`, `ISODD` and `INDIRECT`. `INDIRECT` resolves as a reference as well as a value, so it works as the argument to `SUM`, `INDEX` and the other range-taking functions; it accepts only a reference-shaped string, so the text cannot smuggle in another call. On the 165-workbook calculation benchmark this takes `sheetpedia-9b13ef8593b1` from 0 to 2886 of 2916 cells and `sheetpedia-58e0d15358b7` from 0 to 28 of 28, moving the corpus from 68.4866% to 70.8031% of cells matching Excel with no workbook regressing.
- 9a6cdd5: Restore persisted collaboration snapshots after the recalculation performed on open. Recalculation dirties saved formula caches without creating an authored edit; snapshot adoption now uses the authority's existing pristine-state check, while real local edits and incompatible workbook structures still prevent replacement.
- d1c6a55: xlsx: recalculation now supports the trigonometric family (`SIN`, `COS`, `TAN`, `SINH`, `COSH`, `ASIN`, `ACOS`, `ATAN`, `ATAN2`, `DEGREES`, `RADIANS`), linear-fit statistics (`CORREL`, `SLOPE`, `INTERCEPT`, `COVARIANCE.P`, `COVARIANCE.S`), order statistics (`PERCENTILE.INC`, `QUARTILE.INC`), week numbering (`WEEKNUM` including the ISO type 21, `ISOWEEKNUM`) and delimiter splitting (`TEXTBEFORE`, `TEXTAFTER` with instance counting from either end, case-insensitive matching and an if-not-found fallback). On the 165-workbook calculation benchmark this takes `sheetpedia-441e60078db9` from 306 to 879 of 918 cells and `sheetpedia-8dc99d663d40` from 12 to 380 of 636, moving the corpus from 86.4906% to 87.3023% of cells matching Excel with no workbook regressing.
- b56efb3: xlsx: a sheet declaring `sheetFormatPr/@zeroHeight` now hides the rows that carry no height of their own, as Excel does. The attribute was read nowhere, so those rows took a fitted or default height and rendered visible.

## 0.1.0

### Minor Changes

- fb916eb: Release the DOCX and XLSX Python bindings as 0.1.0. They wrap the 0.2.0 engines and now share a version line with the PPTX binding.

### Patch Changes

- 5069ad2: Keep accepted spreadsheet proposals undoable in collaborative sessions, preserve pending proposals through remote edits, and require a refreshed review when calculated previews change. Reject document suggestions that overlap partially tracked text. Existing public signatures and wire fields remain unchanged.
- 5798031: Load Excel shared formulas by expanding followers into plain cells with correct absolute and relative references, so they evaluate and save with correct values. Shared-formula markup is not written back: an edited sheet writes each follower as its own formula.
- 13016f2: Support whole-column formula references such as `VLOOKUP(..., S:V, ...)` with limits: narrow hits evaluate without materialising the column, but wide aggregates such as `SUM(A:XFD)` return `#NUM!`, and every lookup miss scans the full column height against the shared per-recalculation budget, so a workbook with many misses can turn later formulas `#NUM!`.

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
