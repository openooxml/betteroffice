# betteroffice-xlsx

The typed Rust API for opening, editing, collaborating, calculating, rendering,
and saving XLSX workbooks. Authored sheet state is backed by Yrs and eagerly
materialized into the calculation, rendering, and OOXML pipelines.

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
# Ok::<(), betteroffice_xlsx::Error>(())
```

Saving preserves the source package: parts and sheets an edit did not touch are
copied through byte for byte, and only what changed is reserialized.
The API is experimental and may change before `0.1.0`.

## Collaboration

Collaborative workbooks use explicit, browser-safe client IDs. The host must
assign a unique ID to every runtime replica of a workbook; Yrs cannot detect a
duplicate ID after both replicas begin authoring. Exchange a state vector,
encode the missing v1 update, and apply it on the peer:

```rust
use betteroffice_xlsx::{CalculationOptions, Workbook};

let mut left = Workbook::open_collaborative(&xlsx_bytes, 101)?;
let mut right = Workbook::open_collaborative(&xlsx_bytes, 202)?;

let update = left.encode_diff_v1(&right.encode_state_vector_v1())?;
right.apply_update_v1(&update, CalculationOptions::default())?;
# Ok::<(), betteroffice_xlsx::Error>(())
```

Incoming updates are limited to 64 MiB, state vectors are limited to 65,536
client entries, and at most 4,096 unresolved updates are retained. Updates are
staged on an independent Yrs document before they can change the live workbook.
Frozen collaborative structure includes the logical identities of sheet and
nested shared maps, so replacing those maps is rejected even when their visible
contents match.

Collaboration accepts only the nonstructural schema emitted by this library.
Use authenticated, authorized transports: byte and collection limits reduce
resource exposure but do not make arbitrary hostile Yrs `Any` payloads a safe
input sandbox.

Cell formats are stored in a content-addressed `xlsx:cell-formats` Yrs map, and
per-sheet style maps reference those format keys. Concurrent style creation is
therefore independent of local style-table indices. Collaborative undo and redo
use a Yrs undo manager that tracks only local user-origin transactions; remote
updates and accepted agent proposals are not added to local history.

## Charts

A classic `c:chartSpace` reached from a worksheet or chartsheet drawing is
modelled: every `c:f` reference, the `c:numCache` and `c:strCache` beside it,
and the `xdr:` anchor that pins it. Structural edits move all three together, or
are refused.

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

## Support Matrix

| Capability | Standalone | Collaborative |
| --- | --- | --- |
| Cell content and formatting | Yes | Yes |
| Column widths and row heights | Yes | Yes |
| Formula recalculation and cached save values | Yes | Yes, local projection only |
| Row/column insert and delete | Yes | No |
| Merge and unmerge | Yes | No |
| Add, remove, rename, and restore sheets | Yes | No |
| Chart references and anchors follow structural edits | Yes | Yes |
| Undo/redo | Yes | Yes, local user origin only |
| Agent proposals | Yes | Yes; acceptance is not locally undoable |
| Yrs v1 vectors, diffs, updates, and observers | Encode/observe only | Yes |
