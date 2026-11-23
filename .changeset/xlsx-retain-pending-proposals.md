---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-react': patch
'@betteroffice/xlsx-i18n': patch
'@betteroffice/rust-crates': minor
---

Pending proposals now survive structural edits and undo/redo by default. Targets and formulas rebase; optional `ProposalCell.stale` reports drift, and optional `Proposal.forceBlocked` hides force with a translated reason when deleted targets or unrewritable formulas prevent acceptance. Rust `ProposedEdit` literals must set `stale`.
