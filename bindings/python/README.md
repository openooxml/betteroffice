# betteroffice-xlsx

Read, recalculate, render, and write XLSX workbooks from Python. The engine is
the Rust [BetterOffice](https://betteroffice.dev) XLSX core, compiled into the
wheel — no Excel, no LibreOffice subprocess, no COM.

```bash
pip install betteroffice-xlsx
```

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

`Workbook.open` starts from the values the authoring application cached, and a
cell you have not touched keeps that cached value — including if it was stale
when the file was written. Use `open_recalculated`, or call `recalculate()`,
when you need every formula evaluated by this engine rather than trusted from
the file.

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
left = Workbook.open_collaborative(data)
right = Workbook.open_collaborative(data)
print(left.client_id, right.client_id)

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
edit carries the before and after text a reviewer would compare:

```python
proposal = wb.propose("copilot", [("Sheet1", "H1", "=SUM(B3:B10)")], note="add a total")

for edit in proposal.edits:
    print(edit.address, edit.before, "->", edit.after)   # H1  ''  ->  3600

wb.accept_proposal(proposal.id)     # or wb.reject_proposal(proposal.id)
```

Nothing is written until the proposal is accepted, and accepting applies it as
a single undo step.

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

`openpyxl` is a far broader library and covers plenty this does not. If you need
formulas evaluated or a sheet rasterized, that is the gap this fills.

## API

| | |
| --- | --- |
| `Workbook.open(data)` | open from `bytes` |
| `Workbook.open_path(path)` | open from a path |
| `Workbook.open_recalculated(data)` | open and evaluate every formula |
| `wb.recalculate()` | re-evaluate; returns a `Calculation` summary |
| `wb.sheet_names` / `wb.sheet_count` | sheet metadata |
| `wb[key]` / `wb.sheet(key)` | a `Sheet` by name or index |
| `sheet[addr]` | cell value — see the note below on when it is recalculated |
| `sheet[addr] = value` | set from what a user would type |
| `wb.set_many(sheet, edits)` | write many cells as one undo step |
| `wb.undo()` / `wb.redo()` | walk local history |
| `wb.propose(...)` / `accept_proposal` / `reject_proposal` | staged agent edits |
| `wb.set_style(...)` / `set_number_format(...)` | formatting over a range |
| `wb.open_collaborative(...)` / `diff` / `apply_update` | Yrs replicas |
| `wb.active_sheet` / `set_active_sheet(...)` | read or persist the active tab |
| `sheet.formula(addr)` | source formula, or `None` |
| `wb.render_png(sheet, ...)` | render to PNG |
| `wb.save()` / `wb.save_path(path)` | serialize to XLSX |

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
`str`. `date`, `datetime`, `time`, and `timedelta` raise `TypeError` for now:
converting them needs the workbook's date system, which is not exposed yet, and
stringifying them would write text that only looks like a date. Pass the Excel
serial number as a float if you need a date today.

Strings are interpreted the way Excel interprets typed input: a leading `=` is a
formula, `TRUE`/`FALSE` become booleans, and numeric text becomes a number.
Prefix with an apostrophe to force text.

```python
sheet["A1"] = "'=1+1"   # the text "=1+1"
sheet["A2"] = "=1+1"    # the formula, evaluating to 2.0
```

Mutating calls return a `Mutation` — truthy when something changed, with
`changed` listing the cells the engine *recalculated* as a result. A cell you
wrote directly is not itself a recalculation, so it will not always appear
there.

Errors raise `XlsxError` or a more specific subclass. Invalid peer updates,
broken local collaboration state, stale proposals, and collaboration-only
operations use `InvalidUpdateError`, `CollaborativeStateError`,
`StaleProposalError`, and `NotCollaborativeError`. `StaleProposalError.cells`
lists the changed A1 addresses, and an unknown proposal ID raises `KeyError`.

## Status

`0.0.x`, and the API may change before `0.1.0`. `save` regenerates the package
from the features the model represents; package parts the model does not cover
are not retained, so this is not a round-trip-preserving editor for arbitrary
workbooks.

Wheels are built for Linux (x86_64, aarch64), macOS (arm64, x86_64), and Windows
(x86_64) against the stable ABI for CPython 3.9 and up.

The extension embeds Carlito, a Calibri-metric-compatible face used to measure
and draw cell text, under the SIL Open Font License. Its license travels with
the wheel — see `THIRD-PARTY-NOTICES.md` and `licenses/Carlito-OFL.txt`. The
package's own code is Apache-2.0.

## Links

- [BetterOffice](https://betteroffice.dev) — the project
- [Documentation](https://docs.betteroffice.dev)
- [Source](https://github.com/openooxml/betteroffice) — `bindings/python`
- [betteroffice-xlsx on crates.io](https://crates.io/crates/betteroffice-xlsx) — the engine this wraps

Apache-2.0.
