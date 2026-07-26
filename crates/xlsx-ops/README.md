# betteroffice-xlsx-ops

The invertible op log behind undo, redo, and agent proposals. Applying an op
returns its inverse, so history is replay rather than a stack of snapshots.

- `Op` / `Transaction` / `CellState` — the operation vocabulary and the cell
  state an op reads and writes
- `apply` / `apply_ops` — application, each returning an `InvertedOp`
- `UndoStack` — history built from those inverses
- `Provenance` — who authored an op, which is what lets a human and an agent
  edit the same sheet distinguishably
- `Proposal` / `ProposalSet` / `ProposedEdit` / `ProposalGhost` — staged edits
  an agent proposes and a human accepts or rejects, with per-cell before/after
  ghosts for preview
- `remap` — address remapping, so pending ops and proposals survive row and
  column insertion and deletion
- `input` / `parse_input` — parsing what a user typed into a cell state
- `formatting` — style patches, border presets, and number-format mutations

Used by [betteroffice-xlsx](https://crates.io/crates/betteroffice-xlsx).

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
