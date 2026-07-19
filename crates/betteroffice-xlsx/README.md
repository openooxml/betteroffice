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

Saving regenerates the package from the XLSX features represented by the model;
unsupported package parts are not retained.
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

Undo and redo are intentionally unavailable in collaborative mode until they
can be implemented with a Yrs-aware undo manager. Collaborative edits and
accepted proposals are not added to the standalone inverse-op history, and
`undo`/`redo` return `Error::CollaborativeUndoUnavailable` without mutation.

## Support Matrix

| Capability | Standalone | Collaborative |
| --- | --- | --- |
| Cell content and style | Yes | Yes |
| Column widths and row heights | Yes | Yes |
| Formula recalculation and cached save values | Yes | Yes, local projection only |
| Row/column insert and delete | Yes | No |
| Merge and unmerge | Yes | No |
| Add, remove, rename, and restore sheets | Yes | No |
| Undo/redo | Yes | No |
| Agent proposals | Yes | Yes; acceptance is not locally undoable |
| Yrs v1 vectors, diffs, updates, and observers | Encode/observe only | Yes |
