# betteroffice-vsdx-edit

The collaborative VSDX diagram model, backed by Yrs.

`DiagramSession` stores page ordering, pages, ShapeSheet formulas, and stories
as independently shared types. Formula strings are authoritative: cells never
collapse to evaluated values in collaboration state.

State vectors, diffs, and updates use Yrs v1 and work with any Yjs sync-v1
transport.

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
