# betteroffice-xlsx

Read, recalculate, render, and write XLSX workbooks from Python. Where
`openpyxl` hands back a formula's source text or whatever value the authoring
application happened to cache, this evaluates the formula and can rasterize the
sheet: the Rust [BetterOffice](https://betteroffice.dev) XLSX core is compiled
into the wheel — no Excel, no LibreOffice subprocess, no COM.

```bash
pip install betteroffice-xlsx
```

The distribution is hyphenated, the module is not: `import betteroffice_xlsx`.

## Formulas actually calculate

```python
from betteroffice_xlsx import Workbook

wb = Workbook.open_path("budget.xlsx")

sheet = wb["Sheet1"]
sheet["B1"] = 10
sheet["B2"] = 32
sheet["B3"] = "=SUM(B1:B2)"

print(sheet["B3"])            # 42.0   <- computed here, not read from a cache
print(sheet.formula("B3"))    # 'SUM(B1:B2)'

wb.save_path("budget-out.xlsx")
```

Value and formula are separate accessors on purpose: `sheet["B3"]` is the value,
`sheet.formula("B3")` is the source text. Writing a cell recalculates its
dependents, so the value above is computed here rather than read back from the
file.

`Workbook.open` starts from cached values. Edits recalculate dependent formulas;
`open_recalculated` and `recalculate()` request a full recalculation.

## Render a sheet to PNG

```python
png = wb.render_png("Sheet1", scale=2.0, range="A1:H40")
png.write("preview.png")
print(png.width, png.height)
```

Rendering is the same grid layout and display list the browser editor uses, so
server-side output matches what the web canvas paints.

Opening, recalculating, rendering, and saving release the GIL, so they run in
parallel across threads instead of serializing your workers.

## Collaboration

Every workbook opened with `open_collaborative` is a Yrs replica. The binding
exposes the byte-level primitives rather than a transport, so it drops into a
WebSocket server, a queue, or a test harness without committing you to asyncio:

```python
data = wb.save()
left = Workbook.open_collaborative(data)
right = Workbook.open_collaborative(data)

left["Sheet1"]["B3"] = 1000
right.apply_update(left.diff(right.state_vector()))   # right now agrees

joiner = Workbook.open_collaborative(data)
joiner.apply_update(left.state_as_update())           # catch up from nothing
```

The binding generates a client ID when it is omitted and exposes the chosen ID
through the read-only `client_id` property. A server may pass a deterministic
`client_id` explicitly, but it must be unique among connected peers because Yrs
cannot detect duplicates once two replicas have started authoring. Collaboration
byte inputs accept `bytes`, `bytearray`, and `memoryview`.

## Undo, redo, and batches

```python
wb.set_many("Sheet1", {"H1": 10, "H2": 20, "H3": "=H1+H2"})   # one undo step
wb.undo()
wb.redo()
wb.history()          # History(undo_depth=1, redo_depth=0)
```

Undo covers this replica's own edits. Updates applied from a peer are not in
local history, so undo will not revert someone else's work.

## Agent proposals

An agent can stage edits for a human instead of applying them. Each proposed
edit carries the display text a reviewer would compare — `before` and `after`
are what the cell shows, as strings, while `input` is what would be written:

```python
proposal = wb.propose("copilot", [("Sheet1", "H1", "=B3*2")], note="double the total")

for edit in proposal.edits:
    print(edit.address, repr(edit.before), "->", repr(edit.after))   # H1 '' -> '84'

wb.accept_proposal(proposal.id)     # or wb.reject_proposal(proposal.id)
```

In standalone workbooks, accepting a proposal records one undo step.
`proposals()` lists pending proposals.

A proposal goes stale when one of its target cells changes after it was staged.
`accept_proposal` then raises `StaleProposalError`, whose `cells` names the
addresses that moved underneath it — re-propose against the new values, or
apply it anyway with `force=True`:

```python
from betteroffice_xlsx import StaleProposalError

proposal = wb.propose("copilot", [("Sheet1", "H1", "=B3*3")])
sheet["H1"] = 0
try:
    wb.accept_proposal(proposal.id)
except StaleProposalError as stale:
    print("changed underneath:", stale.cells)   # ['H1']
    wb.accept_proposal(proposal.id, force=True)
```

A changed formula dependency can also make acceptance stale. In that case,
read `proposals()` again to review the refreshed preview before accepting.
Pending proposals remain local to this workbook session; ordinary peer edits
preserve them and acceptance still checks their targets.

An unknown proposal ID raises `KeyError`; `reject_proposal` returns `False`
instead when there is nothing left to reject.

## Version-checked edit batches

`read_cells` returns values, formulas and display text with the workbook
version they were read at. `apply_edits` applies a batch against that version as
one recalculated undo step, or returns a refusal with nothing changed. Requests
and results are the camelCase dictionaries every binding shares, typed in
`betteroffice_xlsx.edits`:

```python
b3 = {"sheetId": "sheet:0", "range": {"kind": "a1", "a1": "B3"}}
read = wb.read_cells({"ranges": [b3]})
result = wb.apply_edits({
    "expectVersion": read["version"],
    "calculation": {"nowSerial": 45658.5},
    "steps": [
        {"op": "setCellInputs", "target": b3, "inputs": [["120"]],
         "expect": {"cells": [[{"displayText": read["ranges"][0]["cells"][0][0]["displayText"]}]]}},
        {"op": "patchStyle", "target": b3, "patch": {"bold": True}},
    ],
})
if not result["ok"]:
    print(result["failure"]["code"])   # "stale-version", "content-mismatch", ...
```

Steps set inputs (parsed like `set`), formulas (source without `=`), number
formats and styles. `validate_edits` stages a batch without changing anything,
`find_text` searches display text exactly, and `history: "none"` keeps a batch
out of undo. A malformed request raises `ValueError`.

## Structured export

`export_structured` returns sparse cells, sheet metadata and diagnostics with
the version they were read at; `export_markdown` renders the same read as
bounded Markdown grids with `<!-- xlsx-export:N -->` markers. Neither
recalculates: formula results are the stored values.

```python
result = wb.export_structured(scope=[{"sheet": 0, "range": "A1:D20"}])
if result["ok"]:
    for cell in result["content"]["sheets"][0]["cells"]:
        print(cell["anchor"]["a1"], cell["value"], cell["formula"], cell["displayText"])

from betteroffice_xlsx import export_xlsx_markdown
print(export_xlsx_markdown(data, markdown_options={"maxRows": 50})["markdown"])
```

Anchors carry `sheet: {"sheetId", "index", "name"}`: a cell or range anchor's
`{"sheetId": anchor["sheet"]["sheetId"], "range": {"kind": "a1", "a1": anchor["a1"]}}`
is its `apply_edits` target at the exported version (`sheet:{index}` ids for
bytes). Hidden sheets, rows, columns and names are excluded unless asked for
(`include_hidden_sheets=True`, ...). Comments, rich-text runs, charts and
pictures are diagnosed or exported as placeholders. `max_cells` and `max_bytes`
stop at a complete record with `truncated`. Unusable scopes come back as
`{"ok": False, ...}`; malformed options raise `ValueError`.

## Formatting

```python
wb.set_number_format("Sheet1", "B3:B10", "#,##0.00")
wb.set_style("Sheet1", "A1:D1", bold=True, fill_color="#eeeeee",
             horizontal_alignment="center")
```

`set_number_format` takes `automatic`, `text`, `number`, `percent`,
`scientific`, `currency`, `date`, `time`, or a custom pattern.

## Reading without recalculating

`Workbook.open` keeps whatever values the file already carried. Use
`open_recalculated` to evaluate everything up front, or call `recalculate()`
later:

```python
wb = Workbook.open_recalculated(open("report.xlsx", "rb").read())

summary = wb.recalculate()
print(summary.changed, summary.cycles)
```

## Compared with openpyxl

| | `openpyxl` | `betteroffice-xlsx` |
| --- | --- | --- |
| Read cell values | yes | yes |
| Evaluate formulas | no — returns the formula string, or a stale cached value | yes |
| Render to an image | no | yes, PNG |
| Engine | pure Python | Rust, compiled |

Use `betteroffice-xlsx` for formula evaluation, PNG rendering and collaborative
editing.

## API

| | |
| --- | --- |
| `Workbook.open(data)` | open from `bytes` |
| `Workbook.open_path(path)` | open from a path |
| `Workbook.open_recalculated(data)` | open and evaluate every formula |
| `Workbook.open_collaborative(data)` | open a Yrs replica — on the class, like the other openers |
| `wb.recalculate()` | re-evaluate; returns a `Calculation` summary |
| `wb.last_calculation()` | the `Calculation` the most recent one produced |
| `wb.sheet_names` / `wb.sheet_count` | sheet metadata |
| `wb[key]` / `wb.sheet(key)` | a `Sheet` by name or index |
| `wb.sheet_index(key)` | resolve a name or index to an index |
| `sheet[addr]` | cell value — see the note below on when it is recalculated |
| `sheet[addr] = value` | set from what a user would type |
| `sheet.formula(addr)` | source formula, or `None` |
| `wb.value(sheet, addr)` / `wb.formula(sheet, addr)` | the same two reads without a `Sheet` |
| `wb.set(sheet, addr, value)` / `wb.set_many(sheet, edits)` | write one cell, or many as one undo step |
| `wb.merged_ranges(sheet, range)` | merged regions overlapping a range |
| `wb.undo()` / `wb.redo()` | walk local history |
| `wb.can_undo` / `wb.can_redo` / `wb.history()` | what history is available |
| `wb.propose(...)` / `proposals()` / `accept_proposal` / `reject_proposal` | staged agent edits |
| `wb.version()` / `read_cells(...)` / `find_text(...)` | versioned reads |
| `wb.validate_edits(...)` / `apply_edits(...)` | version-checked edit batches |
| `wb.export_structured(...)` / `export_markdown(...)` | anchored JSON and Markdown export, never recalculated |
| `export_xlsx_structured(data)` / `export_xlsx_markdown(data)` / `render_xlsx_markdown(content)` | the same from bytes or content |
| `wb.set_style(...)` / `set_number_format(...)` | formatting over a range |
| `wb.diff(sv)` / `apply_update(u)` / `state_vector()` / `state_as_update()` | exchange Yrs updates |
| `wb.client_id` / `wb.is_collaborative` | which kind of workbook you are holding |
| `wb.active_sheet` / `set_active_sheet(...)` | read or persist the active tab |
| `wb.render_png(sheet, ...)` | render to PNG |
| `wb.save()` / `wb.save_path(path)` | serialize to XLSX |

The following calculation calls — `open_recalculated`,
`open_collaborative`, `recalculate`, `set`, `set_many`, `undo`, `redo`,
`apply_update`, `propose`, `accept_proposal`, `set_number_format`, and
`set_style` — take a keyword-only `now_serial`. It is the clock `TODAY()` and
`NOW()` read, as an Excel serial number. The engine has none of its own, so
both return `#VALUE!` unless you pass one:

```python
wb.recalculate(now_serial=45658.5)   # 2025-01-01, midday
```

Cell values come back as `None`, `float`, `str`, `bool`, or `CellError`.
Numbers are `f64` in the engine, so they arrive as `float` and are not narrowed
to `int`. Errors are a `CellError` instance rather than a string, so `#DIV/0!`
as a value is distinguishable from a cell containing that text. `CellError`
compares equal to its code, and hashes like it, so it works as a dict key:

```python
if sheet["D3"] == "#DIV/0!":
    ...
```

Writing accepts `None` (clears the cell), `bool`, `int`, `float`, `Decimal`, and
`str`. `date`, `datetime`, `time` and `timedelta` raise `TypeError`; pass Excel
serials as floats.

General cells interpret strings like Excel: a leading `=` is a formula,
`TRUE`/`FALSE` become booleans, and numeric text becomes a number. Text (`@`)
cells preserve input as text. Prefix with an apostrophe to force text in any format.

```python
sheet["A1"] = "'=1+1"   # the text "=1+1"
sheet["A2"] = "=1+1"    # the formula, evaluating to 2.0
```

Mutating calls return a truthy-on-change `Mutation`. `changed` lists
recalculated cells; `cycles` and `limited` identify calculation diagnostics.
Addresses are A1 on the active sheet and `Sheet!A1` elsewhere.

Errors raise `XlsxError` or a more specific subclass: `ParseError`,
`RangeError`, `RenderError`, `InvalidUpdateError`, `CollaborativeStateError`,
`StaleProposalError`, `NotCollaborativeError`. Invalid peer updates, broken
local collaboration state, stale proposals, and collaboration-only operations
are the last four in that order. `StaleProposalError.cells` lists the changed A1
addresses and `targets` their sheets and coordinates, and an unknown proposal ID
raises `KeyError`.

## Status

Pre-1.0: the API may change between minor versions.

`save` keeps the parts the model does not represent — charts, drawings, pivot
tables, comments, macros, custom XML, and their relationships — rather than
regenerating the package from the modeled features, and sheets you did not
touch are copied through byte for byte. The stylesheet is left alone unless
styles actually change.

Saving serializes edited worksheet state and retains source range coordinates.
Collaboration fingerprints the modeled workbook.

Wheels are built for Linux (x86_64, aarch64), macOS (arm64, x86_64), and Windows
(x86_64) against the stable ABI for CPython 3.9 and up.

The extension embeds Carlito, a Calibri-metric-compatible face used to measure
and draw cell text, under the SIL Open Font License. Its license travels with
the wheel — see `THIRD-PARTY-NOTICES.md` and `licenses/Carlito-OFL.txt`. The
package's own code is Apache-2.0.

## Links

- [BetterOffice](https://betteroffice.dev) — the project
- [Documentation](https://docs.betteroffice.dev/docs/python)
- [Source](https://github.com/openooxml/betteroffice) — `bindings/python-xlsx`
- [betteroffice-xlsx on crates.io](https://crates.io/crates/betteroffice-xlsx) — the engine this wraps

Apache-2.0.
