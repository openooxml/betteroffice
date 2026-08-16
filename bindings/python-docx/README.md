# betteroffice-docx

Read, edit, lay out, and rasterize DOCX documents from Python. `python-docx`
reads a document and writes one back; this also *paginates* it — page boxes, a
display list, PNG pages — because the Rust
[BetterOffice](https://betteroffice.dev) DOCX core is compiled into the wheel:
no Word, no LibreOffice subprocess, no COM.

```bash
pip install betteroffice-docx
```

The distribution is hyphenated, the module is not: `import betteroffice_docx`.

## Read a document

```python
from betteroffice_docx import Document

document = Document.open_path("report.docx")

print(document.structure())          # Structure(body_paragraphs=42, body_tables=3, sections=2)
for paragraph in document:
    print(paragraph.id, paragraph.style, repr(paragraph.text))

for table in document.tables():
    for row in table.rows:
        print([cell.text for cell in row.cells])
```

`paragraphs()` walks the body in document order and descends into table cells
and content controls, so a cell paragraph is reachable both ways.
`document[key]` and `document.paragraph(key)` take either a `w14:paraId` or a
body index. Everything a read returns is a value, not a live view — read again
after an edit.

Page geometry is in twips — 1440 to the inch. The module exports
`TWIPS_PER_INCH` and `TWIPS_PER_POINT`.

```python
section = document.sections()[0]
print(section.page_width, section.page_height, section.margin_left)
print(document.headers()[0].text)
```

## Edit text

```python
edit = document.replace_text("11111111", "Edited from Python")
print(edit.para_id, edit.start, edit.end)

document.save_path("report-edited.docx")
```

`replace_text` rewrites one paragraph and keeps its style, alignment, and the
run formatting it already had. The engine rebuilds the paragraph from a single
run, so a paragraph that mixes runs — half bold, a hyperlink, a field — raises
`UnsupportedEditError` rather than flattening the formatting you did not ask it
to touch. An unknown `w14:paraId` raises `KeyError`.

Only paragraphs Word stamped with a `w14:paraId` can be addressed:
`document.paragraph_ids` reports `None` for the rest.

## Write

Unlike the PPTX binding, edits reach the file: `save()` serializes the edited
model, and reopening the result gives the edited text back.

```python
document = Document.open(data)
document.replace_text(document.paragraph_ids[0], "New first line")
reopened = Document.open(document.save())
reopened.paragraph(0).text          # 'New first line'
```

Saving is deterministic. The engine has no clock, so timestamps come from
`document.timestamp` — the epoch until you set one — and the same input plus the
same edits produce the same bytes. `save(now=..., update_modified_date=True,
modified_by=...)` overrides that for one call.

The container is rebuilt rather than patched, so output is not byte-identical to
the source even with no edits; the parts the model retained survive unchanged.

## Lay a document out

Layout is a two-stage contract. Something else measures text — the browser, or
`ooxml-text` — and the engine paginates the measured blocks and compiles them
into a display list:

```python
layout = document.layout({"measured": measured_blocks, "options": {...}})
print(len(layout), layout.pages)
layout.write("layout.json")

pages = layout.display_list
print(len(pages), pages.primitives)
```

`layout()` takes the envelope as a `dict` or as a JSON string, and returns the
page boxes (`layout.json`, `layout.to_dict()`) beside the display list that
paints them.

## Rasterize

**No font is compiled into the wheel**, so a page with text needs at least one
registered face:

```python
from pathlib import Path

document.register_font("Carlito", Path("Carlito-Regular.ttf").read_bytes())
document.register_font("Carlito", Path("Carlito-Bold.ttf").read_bytes(), bold=True)

png = document.render_png(layout.display_list, 0)
png.write("page-0.png")
print(len(png), png.skipped_images)
```

Text whose family has no chain raises `RenderError` naming the chain it wanted
— ``missing font chain for `calibri|0|0` `` — so a missing face is loud rather
than silently blank.

Images are the opposite: an image reference the backend cannot resolve is
skipped and counted in `png.skipped_images` instead of failing the page. Word
hands out relationship ids per part, so `rId9` in the body and `rId9` in a
header are different images and registration is scoped:

```python
document.register_image("rId9", body_png)
document.register_image("rId9", header_png, scope="header_footer", part="rId7")
document.register_image("rId4", note_png, scope="footnotes")
```

Images the display list already carries as `data:` URLs need no registration.
A page past `MAX_PIXMAP_DIM` per side or `MAX_PIXMAP_PIXELS` in area is refused
before any surface is allocated.

## Compared with python-docx

| | `python-docx` | `betteroffice-docx` |
| --- | --- | --- |
| Read paragraphs, tables, sections | yes | yes |
| Write text back to a file | yes | yes, single-run paragraphs |
| Build a document from scratch | yes | no — it edits what you open |
| Paginate (page boxes, display list) | no | yes |
| Rasterize pages to PNG | no | yes |
| Engine | pure Python | Rust, compiled |

`python-docx` is a far broader authoring library. If what you need is
pagination, page images, or an engine that reads what Word actually wrote, that
is the gap this fills.

## API

| | |
| --- | --- |
| `Document.open(data)` / `open_path(path)` | open from bytes or a path |
| `document.structure()` | paragraph, table, section, and note counts |
| `document.paragraphs()` / `tables()` / `sections()` | body content |
| `document.headers()` / `footers()` | header and footer stories |
| `document[key]` / `document.paragraph(key)` | one paragraph by ID or index |
| `document.paragraph_ids` / `text` | body IDs, and the whole text |
| `document.warnings` / `template_variables` | what the parser found |
| `document.replace_text(para_id, text)` | rewrite one paragraph |
| `document.author` / `origin` / `timestamp` | how an edit is attributed and stamped |
| `document.layout(input)` | paginate a measured envelope |
| `document.register_font` / `register_image` | raster resources |
| `document.render_png(display_list, page)` | rasterize one page |
| `document.save()` / `save_path(path)` | serialize to DOCX |

Errors raise `DocxError` or a more specific subclass: `ParseError`,
`EditError`, `UnsupportedEditError`, `LayoutError`, `RenderError`. An unknown
paragraph ID raises `KeyError`, an out-of-range index `IndexError`, and a bad
argument — an unknown parse limit, an unknown image scope, malformed font bytes
— `ValueError`.

Parser bounds can be tightened for untrusted input:

```python
Document.open(untrusted, limits={"max_paragraphs": 5_000, "max_tables": 500})
```

An unknown limit name raises `ValueError` rather than being ignored.

## Threads

A `Document` is not pinned to a thread: the engine's document type is `Send` and
`Sync`, so opening on one thread and dropping on another is fine. Parsing,
layout, rasterization, and saving release the GIL for their duration, so several
documents genuinely proceed in parallel.

## Status

`0.0.x`, and the API may change before `0.1.0`. Editing covers paragraph text on
plain single-run paragraphs; richer edits land on the Rust facade first.

Wheels are built for Linux (x86_64, aarch64), macOS (arm64, x86_64), and Windows
(x86_64) against the stable ABI for CPython 3.9 and up.

## Links

- [BetterOffice](https://betteroffice.dev) — the project
- [Documentation](https://docs.betteroffice.dev/docs/python)
- [Source](https://github.com/openooxml/betteroffice) — `bindings/python-docx`
- [betteroffice-docx on crates.io](https://crates.io/crates/betteroffice-docx) — the engine this wraps

Apache-2.0.
