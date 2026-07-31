# betteroffice-pptx

Read, edit, and lay out PPTX presentations from Python. The engine is the Rust
[BetterOffice](https://betteroffice.dev) PPTX core, compiled into the wheel —
no PowerPoint, no LibreOffice subprocess, no COM.

Editing is in-memory today: see [Writing](#writing) before you reach for
`save()`.

```bash
pip install betteroffice-pptx
```

## Read a deck

```python
from betteroffice_pptx import Presentation

deck = Presentation.open_path("quarterly.pptx")

for slide in deck:
    print(slide.index, slide.name, repr(slide.text))

shape = deck[0].shapes[0]
print(shape.kind, shape.geometry, shape.x, shape.y, shape.width, shape.height)
```

Geometry is in English Metric Units — 914400 to the inch. The module exports
`EMU_PER_INCH`, `EMU_PER_CENTIMETER`, and `EMU_PER_POINT` so you rarely have to
type the constant.

`snapshot()` returns the whole deck as plain data in one call; `slide(key)`
returns one slide by index or by ID. Both are values, not live views — read
them again after an edit.

## Edit shapes and slides

```python
from betteroffice_pptx import EMU_PER_INCH as INCH

deck = Presentation.open(open("deck.pptx", "rb").read())
slide_id = deck.slide_ids[0]

edit = deck.add_text_box(
    slide_id, x=INCH, y=INCH, width=4 * INCH, height=INCH,
    text="Revenue", bold=True, font_size=32.0,
)
deck.move_shape(slide_id, edit.shape_id, 2 * INCH, INCH)
deck.resize_shape(slide_id, edit.shape_id, 5 * INCH, 2 * INCH)

box = deck.add_shape(slide_id, "roundRect", x=INCH, y=3 * INCH,
                     width=2 * INCH, height=INCH, fill="#2563eb")
deck.set_shape_stroke(slide_id, box.shape_id, color="#111827", width_pt=2.0)
deck.set_shape_adjust(slide_id, box.shape_id, {"adj": 0.25})
```

Every mutating call returns a receipt naming what it touched: `ShapeEdit` has
the new shape's ID and z-order index, `TransformEdit` carries the rect before
and after, and `AdjustEdit` shows the values after the engine clamped them into
their guide's legal range.

An unsupported preset geometry, a non-positive size, or an unknown adjustment
guide raises `ValueError` instead of writing a shape PowerPoint would reject.

## Edit text

Text lives in *stories* — one editable flow per text-bearing shape. Offsets are
UTF-16 code units, and every paragraph ends with a pilcrow occupying one of
them.

```python
story = deck[0].shapes[0].stories[0]
print(story.text, story.length)

deck.insert_text(story.id, 0, "Q3 ", bold=True)
deck.format_text(story.id, 0, 3, color="#dc2626")
deck.insert_paragraph_break(story.id, 3)
removed = deck.delete_text(story.id, 0, 3)
print(removed.text)          # 'Q3 '
```

`format_text` patches only the arguments you pass. A range that crosses a
paragraph boundary raises `RangeError` rather than silently splitting the edit.

## Lay a slide out

```python
deck.register_font("Inter", open("Inter-Regular.ttf", "rb").read())
deck.register_font("Inter", open("Inter-Bold.ttf", "rb").read(), bold=True)

layout = deck.render_slide(0)
print(layout.width, layout.height, len(layout))   # 1280.0 720.0 42
layout.write("slide-0.json")
scene = layout.to_dict()
```

`render_slide` returns the display list — the same drawing contract the browser
editor paints, as JSON. There is no PPTX rasterizer yet, so this is a scene
description rather than pixels; feed it to your own canvas or renderer.

No font is compiled into the wheel, so families you do not register fall back
to the engine's metrics-only path. Register the faces you actually use.

## Collaboration

`open_collaborative` gives this replica a unique client ID, which peers need in
order to converge:

```python
left = Presentation.open_collaborative(data)
right = Presentation.open_collaborative(data)

left.move_shape(0, left[0].shapes[0].id, 100000, 100000)
right.apply_update(left.diff(right.state_vector()))   # right now agrees

joiner = Presentation.open_collaborative(data)
joiner.apply_update(left.state_as_update())           # catch up from nothing
```

A deck from `open` or `open_path` is *not* a replica: it has no client ID of its
own, so two of them would author under the same identity and never converge.
`state_vector`, `state_as_update`, `diff`, and `apply_update` raise
`NotCollaborativeError` on such a deck rather than diverging silently, and
`is_collaborative` says which kind you are holding:

```python
deck = Presentation.open(data)
deck.is_collaborative                  # False
deck.state_vector()                    # NotCollaborativeError
```

The binding generates a client ID when it is omitted and exposes it through the
read-only `client_id` property. Explicit IDs must be unique among connected
peers, because Yrs cannot detect duplicates once two replicas have started
authoring. Byte inputs accept `bytes`, `bytearray`, and `memoryview`, and an
oversized payload is refused before it is copied.

## Undo, redo, and attribution

```python
deck.author = "ana"
deck.move_shape(0, shape_id, 0, 0)
deck.add_undo_barrier()      # the next edit starts a new undo step
deck.undo()
deck.redo()
```

Undo covers this replica's own local edits. Updates applied from a peer are not
in local history, so undo will not revert someone else's work. Consecutive edits
inside half a second coalesce into one step; `add_undo_barrier()` splits them.

Setting `origin` to `"agent"`, `"remote"`, or `"system"` tags edits for
attribution — and takes them out of the local undo stack, which is the point:
an agent's write is not something the user undoes by accident.

## Writing

**Edits do not reach a saved file yet.** The engine serializes the parsed
package, not the edited model, so `save()` on a deck you have edited would
silently drop every change. Rather than hand you a file that quietly lost your
work, `save()` and `save_path()` refuse once anything has been edited:

```python
deck = Presentation.open(data)
deck.save_path("copy.pptx")        # fine: nothing was edited

deck.insert_slide(1)
deck.is_edited                     # True
deck.save()                        # UnsupportedWriteError
```

`UnsupportedWriteError` is its own `PptxError` subclass, so "cannot write edits
yet" is distinguishable from a zip or filesystem failure without matching on a
message, and `is_edited` lets a pipeline branch before it does expensive work
rather than discovering the refusal at the end. Only an edit the engine
*accepted* sets it: an edit that raised leaves the deck saveable.

Reading, laying out, and collaborative editing are fully usable today; treat
this release as read, render, and edit-in-memory. Write-back is the next thing
to land, and the refusal disappears when it does.

## Compared with python-pptx

| | `python-pptx` | `betteroffice-pptx` |
| --- | --- | --- |
| Read shapes and text | yes | yes |
| Write shapes and text back to a file | yes | not yet — see *Writing* |
| Lay slides out (line breaking, text metrics) | no | yes, display list |
| Collaborative editing (CRDT) | no | yes, Yrs |
| Undo/redo | no | yes |
| Engine | pure Python | Rust, compiled |

`python-pptx` is a far broader library and covers plenty this does not —
charts, tables, and templating in particular. If you need slides laid out, or
edits that merge across replicas, that is the gap this fills.

## API

| | |
| --- | --- |
| `Presentation.open(data)` / `open_path(path)` | open from bytes or a path |
| `Presentation.open_collaborative(data)` | open a Yrs replica |
| `deck.snapshot()` | the whole deck as plain data |
| `deck[key]` / `deck.slide(key)` | a `Slide` by index or ID |
| `deck.slide_ids` / `slide_count` / `layouts` | deck metadata |
| `deck.story(id)` | one text flow |
| `deck.media()` | embedded images and other binary parts |
| `insert_slide` / `delete_slide` / `move_slide` | slide order |
| `add_text_box` / `add_shape` / `remove_shape` | shape lifecycle |
| `move_shape` / `resize_shape` | shape geometry |
| `set_shape_fill` / `set_shape_stroke` / `set_shape_adjust` | shape styling |
| `insert_text` / `delete_text` / `format_text` | text editing |
| `insert_paragraph_break` | split a paragraph |
| `register_font` / `render_slide` | layout |
| `diff` / `apply_update` / `state_vector` / `state_as_update` | Yrs replicas |
| `deck.is_collaborative` | whether this deck may exchange updates |
| `deck.is_edited` | whether an accepted edit has made `save` refuse |
| `undo` / `redo` / `add_undo_barrier` / `can_undo` / `can_redo` | history |
| `deck.save()` / `save_path(path)` | serialize to PPTX — see *Writing* |

Errors raise `PptxError` or a more specific subclass: `ParseError`,
`RangeError`, `RenderError`, `InvalidUpdateError`, `CollaborativeStateError`,
`NotCollaborativeError`, `UnsupportedWriteError`.
An unknown slide, shape, or story ID raises `KeyError`; a bad argument — an
unsupported geometry, an out-of-range client ID, an unknown parse limit —
raises `ValueError`.

Parser bounds can be tightened for untrusted input:

```python
Presentation.open(untrusted, limits={"max_shapes": 5_000, "max_runs": 50_000})
```

An unknown limit name raises `ValueError` rather than being ignored.

## Status

`0.0.x`, and the API may change before `0.1.0`. `save` re-zips the parts the
model retained; part bytes survive unchanged but the container is rebuilt, so
output is not byte-identical to the source, and it refuses on an edited deck —
see *Writing*.

A `Presentation` is bound to the thread that opened it — the engine's undo
manager is not `Send`. Give each worker its own `Presentation` rather than
sharing one across a pool. For the same reason this binding holds the GIL for
the duration of a call, unlike `betteroffice-xlsx`.

Wheels are built for Linux (x86_64, aarch64), macOS (arm64, x86_64), and Windows
(x86_64) against the stable ABI for CPython 3.9 and up.

## Links

- [BetterOffice](https://betteroffice.dev) — the project
- [Documentation](https://docs.betteroffice.dev)
- [Source](https://github.com/openooxml/betteroffice) — `bindings/python-pptx`
- [betteroffice-pptx on crates.io](https://crates.io/crates/betteroffice-pptx) — the engine this wraps

Apache-2.0.
