# betteroffice-vsdx

Read Visio VSDX diagrams from Python on the BetterOffice Rust engine.

```python
from betteroffice_vsdx import Diagram

diagram = Diagram.open_path("architecture.vsdx")
for page in diagram.pages:
    for shape in page.shapes:
        print(page.id, shape.id, shape.name, shape.text)
```

Geometry uses Visio's native inches. `Shape.pin_x`, `pin_y`, `width`, and
`height` are inches; cells preserve their Visio formula, value, and unit.

`Diagram.open()` accepts `bytes`, `bytearray`, and `memoryview`. Values returned
by `pages` are snapshots, so open the diagram again after changing its source.
`cells` contains direct shape cells; section rows are not exposed. Geometry
properties read stored numeric values, and `text` does not expand fields or
inherited master text.

## API

| API | Description |
| --- | --- |
| `Diagram.open(data)` / `open_path(path)` | Open bytes or a VSDX file |
| `diagram.pages` | Snapshot pages |
| `Page.shapes` / `connects` | Top-level shapes and glued connections |
| `Shape.children` / `cells` | Group descendants and ShapeSheet cells |

Errors raise `VsdxError`, or `ParseError` for invalid packages and XML.
