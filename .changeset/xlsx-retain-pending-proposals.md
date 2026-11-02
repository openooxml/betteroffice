---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-react': patch
---

Pending proposals now survive structural edits and undo/redo by default. Targets and formulas rebase with cells; deleted targets remain visible with optional `ProposalCell.stale` state and refuse acceptance, including force, until restored.
