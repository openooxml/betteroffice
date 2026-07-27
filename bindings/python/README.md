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

Errors raise `XlsxError`, or one of `ParseError`, `RangeError`, `RenderError`.

## Status

`0.0.x`, and the API may change before `0.1.0`. `save` regenerates the package
from the features the model represents; package parts the model does not cover
are not retained, so this is not a round-trip-preserving editor for arbitrary
workbooks. Collaboration, agent proposals, undo/redo, and the styling APIs exist
in the Rust engine but are not yet exposed here.

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
