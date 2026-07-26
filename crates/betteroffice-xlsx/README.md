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

`save` regenerates the package from the features the model represents.
Package parts the model does not cover are not retained.

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

## Support matrix

| Capability | Standalone | Collaborative |
| --- | --- | --- |
| Cell content and formatting | Yes | Yes |
| Column widths and row heights | Yes | Yes |
| Formula recalculation and cached save values | Yes | Local projection only |
| Row/column insert and delete | Yes | No |
| Merge and unmerge | Yes | No |
| Add, remove, rename, and restore sheets | Yes | No |
| Undo/redo | Yes | Local user origin only |
| Agent proposals | Yes | Yes; acceptance is not locally undoable |
| Yrs v1 vectors, diffs, updates, and observers | Encode/observe only | Yes |

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
