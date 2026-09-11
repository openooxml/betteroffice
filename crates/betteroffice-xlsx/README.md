# betteroffice-xlsx

Open, edit, calculate, render, and save XLSX workbooks from native Rust. Sheet
state lives in Yrs and is materialized eagerly into the calculation, rendering,
and OOXML pipelines.

```rust
use betteroffice_xlsx::{
    CalculationOptions, CellRef, RenderOptions, SheetId, Workbook,
};

let mut workbook = Workbook::open_recalculated(
    &xlsx_bytes,
    CalculationOptions::default(),
)?;

workbook.edit_cell(
    SheetId(0),
    CellRef::parse_a1("A1").unwrap(),
    "42",
    CalculationOptions::default(),
)?;

let png = workbook.render_sheet(SheetId(0), &RenderOptions::default())?;
let saved = workbook.save()?;
```

Saving preserves the source package: parts and sheets an edit did not touch are
copied through byte for byte, and only what changed is reserialized.
`0.0.x`: the API may change before `0.1.0`.

## Collaboration

Every replica needs an explicit client ID, assigned by the host. Yrs cannot
detect a duplicate once two replicas have started authoring. Exchange a state
vector, encode the missing v1 update, apply it on the peer:

```rust
use betteroffice_xlsx::{CalculationOptions, Workbook};

let mut left = Workbook::open_collaborative(&xlsx_bytes, 101)?;
let mut right = Workbook::open_collaborative(&xlsx_bytes, 202)?;

let update = left.encode_diff_v1(&right.encode_state_vector_v1())?;
right.apply_update_v1(&update, CalculationOptions::default())?;
```

Incoming updates cap at 64 MiB, state vectors at 65,536 client entries, and
4,096 unresolved updates. Updates stage on a separate Yrs document before they
can touch the live workbook. Sheet and nested shared maps are frozen by logical
identity, so replacing one is rejected even when its visible contents match.

Only the nonstructural schema this library emits is accepted. Use authenticated,
authorized transports — the byte and collection limits reduce resource exposure,
they do not make hostile Yrs `Any` payloads safe.

Cell formats live in a content-addressed `xlsx:cell-formats` map that per-sheet
style maps reference by key, so concurrent style creation does not depend on
local style-table indices. Undo and redo track local user-origin transactions
only; remote updates and accepted agent proposals stay out of local history.

## Charts

A classic `c:chartSpace` reached from a worksheet or chartsheet drawing is
modelled: every `c:f` reference, the `c:numCache` and `c:strCache` beside it,
and the `xdr:` anchor that pins it. Structural edits move all three together, or
are refused.

A chart on a sheet is a frame, addressed by the drawing part and the anchor
index within it — `ChartRegion::id`, what `move_chart` and `Op::SetChartAnchor`
take. Two anchors may point at one chart part, so the part names no single
frame; each frame hit-tests, moves and saves as itself. An index names a
position, not an object, so `Op::SetChartAnchor` is a compare-and-set: it
carries the chart part and the anchor the frame held when the op was recorded,
and a replay onto a drawing another editor has since added to, reordered or
thinned is refused as `Error::ChartFrameShifted` rather than landing on
whichever frame now sits at that index.

A chart part is refused, and with it every rename, sheet removal and row or
column edit on the workbook, when it is a ChartEx (`cx:chartSpace`) part, when
no sheet claims it, when it carries a pivot source, an external data cache, a
reference outside the chart namespace or a `sqref` extension reference, or when
a cache sits beside a reference this crate cannot resolve — a defined name, a
union, an external book, an area spanning two dimensions, or a multi-level
category cache. Part types are resolved the way OPC resolves them, an
`Override` first and the `Default` for the extension after, so a chart typed by
extension alone is found too.

A save is refused when a moved reference names a sheet the workbook does not
hold, when a string cache covers cells that are not text, when a cache would
hold more points than a chart can carry, or when the cache carries content this
crate does not model, such as an `extLst` or a per-point `formatCode`.

Charts are preserved, never created. A model carrying charts with no source
package is refused, as is a chart that appeared on or vanished from a sheet
between open and save, and an anchor change beyond the row and column a save
writes back.

## Support matrix

| Capability | Standalone | Collaborative |
| --- | --- | --- |
| Cell content and formatting | Yes | Yes |
| Column widths and row heights | Yes | Yes |
| Formula recalculation and cached save values | Yes | Local projection only |
| Row/column insert and delete | Yes | No |
| Merge and unmerge | Yes | No |
| Add, remove, rename, and restore sheets | Yes | No |
| Chart references and anchors follow structural edits | Yes | Yes |
| Undo/redo | Yes | Local user origin only |
| Agent proposals | Yes | Yes; acceptance is not locally undoable |
| Yrs v1 vectors, diffs, updates, and observers | Encode/observe only | Yes |

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
