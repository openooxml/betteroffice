# betteroffice-pptx

Read, edit, and lay out PPTX presentations from Python. `python-pptx` reads a
deck and writes one back; this also *lays slides out* — line breaking, text
metrics, a display list — and merges edits across replicas, because the Rust
[BetterOffice](https://betteroffice.dev) PPTX core is compiled into the wheel:
no PowerPoint, no LibreOffice subprocess, no COM.

```bash
pip install betteroffice-pptx
```

The distribution is hyphenated, the module is not: `import betteroffice_pptx`.

## Read a deck

```python
from betteroffice_pptx import Presentation

deck = Presentation.open_path("quarterly.pptx")

for slide in deck:
    print(slide.index, slide.name, repr(slide.text))

shape = next((s for slide in deck for s in slide.shapes), None)
if shape is not None:
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
story = next(
    (s for slide in deck for shape in slide.shapes for s in shape.stories), None
)
if story is not None:
    print(story.text, story.length)

    deck.insert_text(story.id, 0, "Q3 ", bold=True)
    deck.format_text(story.id, 0, 3, color="#dc2626")
    deck.insert_paragraph_break(story.id, 3)
    removed = deck.delete_text(story.id, 0, 3)
    print(removed.text)      # 'Q3 '
```

A shape with no text has no story, so a deck of pictures alone yields none.
`deck.story(id)` looks one up directly and raises `KeyError` if it is gone.

`format_text` patches only the arguments you pass, and a range spanning several
paragraphs styles each of them as a single undoable edit. `delete_text` is the
strict one: a range crossing a paragraph boundary raises `RangeError` rather
than silently swallowing the break.

## Lay a slide out

**No font is compiled into the wheel**, so laying out a slide that has text
needs at least one registered face. Before that, `render_slide` raises:

```python
deck.render_slide(0)
# RenderError: no font has been registered for slide text
```

Register the faces the deck uses — one call per family, weight, and slant — and
it lays out:

```python
from pathlib import Path

deck.register_font("Inter", Path("Inter-Regular.ttf").read_bytes())
deck.register_font("Inter", Path("Inter-Bold.ttf").read_bytes(), bold=True)

layout = deck.render_slide(0)
print(layout.width, layout.height, len(layout))   # 1280.0 720.0 42
layout.write("slide-0.json")
scene = layout.to_dict()
```

Once at least one face exists nothing raises again: a family the deck names but
you never registered resolves to the same family at regular weight, and failing
that to the first face you registered at all. One registration therefore renders
every slide — in that one typeface, at its metrics. Register the real faces when
line breaking has to match what PowerPoint would do.

`render_slide` returns the display list — the same drawing contract the browser
editor paints, as JSON. There is no PPTX rasterizer yet, so this is a scene
description rather than pixels; feed it to your own canvas or renderer.

## Collaboration

`open_collaborative` gives this replica a unique client ID, which peers need in
order to converge:

```python
left = Presentation.open_collaborative(data)
right = Presentation.open_collaborative(data)

left.add_text_box(0, x=INCH, y=INCH, width=4 * INCH, height=INCH, text="Q3")
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
edit = deck.add_text_box(0, x=INCH, y=INCH, width=INCH, height=INCH, text="Q3")
deck.move_shape(0, edit.shape_id, 0, 0)
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

`save()` and `save_path()` serialize the deck with every accepted edit
applied. Slides you did not touch keep their exact source part bytes; edited
slides are patched at the XML level, so unmodeled markup survives:

```python
deck = Presentation.open(data)
deck.insert_slide(1)
deck.is_edited                     # True
deck.save_path("copy.pptx")        # edits included

reopened = Presentation.open_path("copy.pptx")
reopened.slide_count               # one more than the source
```

`is_edited` reports whether the engine has accepted an edit since the deck was
opened. Only an edit the engine *accepted* sets it: an edit that raised leaves
the flag untouched.

## Compared with python-pptx

| | `python-pptx` | `betteroffice-pptx` |
| --- | --- | --- |
| Read shapes and text | yes | yes |
| Write shapes and text back to a file | yes | yes — see *Writing* |
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
| `deck.width_emu` / `height_emu` | slide size in EMU |
| `deck.author` / `deck.origin` | who an edit is attributed to, and how |
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
| `deck.is_collaborative` / `deck.client_id` | whether this deck may exchange updates, and as whom |
| `deck.is_edited` | whether the engine has accepted an edit since open |
| `undo` / `redo` / `add_undo_barrier` / `can_undo` / `can_redo` | history |
| `deck.save()` / `save_path(path)` | serialize to PPTX — see *Writing* |

Errors raise `PptxError` or a more specific subclass: `ParseError`,
`RangeError`, `RenderError`, `InvalidUpdateError`, `CollaborativeStateError`,
`NotCollaborativeError`.
An unknown slide, shape, or story ID raises `KeyError`; a bad argument — an
unsupported geometry, an out-of-range client ID, an unknown parse limit —
raises `ValueError`.

Parser bounds can be tightened for untrusted input:

```python
Presentation.open(untrusted, limits={"max_shapes": 5_000, "max_runs": 50_000})
```

An unknown limit name raises `ValueError` rather than being ignored.

## Threads

**A `Presentation` is pinned to the thread that opened it, and must also be
released there.** The engine's undo manager is not `Send`, so the class is
declared `unsendable`, and pyo3 enforces that in two ways worth knowing about:

- **Touching one from another thread raises `pyo3_runtime.PanicException`.**
  That is a direct `BaseException` subclass, so `except Exception` does *not*
  catch it — a worker that guards its work with `except Exception` will die
  anyway.
- **Releasing one on another thread leaks it.** pyo3 skips the Rust destructor
  and writes an *unraisable* `RuntimeError` instead (visible only through
  `sys.unraisablehook`), stranding roughly 1.5 MB per deck. Nothing is raised
  into your code.

The leak is easy to hit by accident, because the release does not have to be an
explicit `del`:

```python
with ThreadPoolExecutor() as pool:
    decks = [f.result() for f in [pool.submit(load, p) for p in paths]]
# every deck was opened on a worker and is now dropped on the main thread
```

The cyclic garbage collector counts too. If a `Presentation` is caught in a
reference cycle — a traceback that reaches it, an object graph that points back
at itself — the collector frees it wherever it happens to run, which may be any
thread. Giving each worker its own `Presentation` therefore is *not* enough on
its own; the deck must also become garbage on its owning thread. Open, use, and
drop each deck inside one thread, and break any cycle holding it before that
thread finishes.

Engine calls hold the GIL for their duration, unlike `betteroffice-xlsx`. Only
the file I/O releases it: `open_path`'s read, and the writes in `save_path`,
`Media.write` and `DisplayList.write`.

## Status

`0.0.x`, and the API may change before `0.1.0`. `save` writes edits back at the
XML level and copies untouched parts through byte for byte; the container is
rebuilt, so output is not byte-identical to the source — see *Writing*.

Wheels are built for Linux (x86_64, aarch64), macOS (arm64, x86_64), and Windows
(x86_64) against the stable ABI for CPython 3.9 and up.

## Links

- [BetterOffice](https://betteroffice.dev) — the project
- [Documentation](https://docs.betteroffice.dev/docs/python)
- [Source](https://github.com/openooxml/betteroffice) — `bindings/python-pptx`
- [betteroffice-pptx on crates.io](https://crates.io/crates/betteroffice-pptx) — the engine this wraps

Apache-2.0.
