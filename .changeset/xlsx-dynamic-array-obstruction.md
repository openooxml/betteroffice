---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: a value typed, pasted, proposed, batched or written by a raw operation into a dynamic array's spill is now kept and makes the array show `#SPILL!`, as in Excel, instead of being overwritten on the next recalculation. Every collaborating replica settles the same way however the edits reach it, undo and redo (local or collaborative) restore the spill, and an edit's result lists the anchor and spilled cells it changed. Clearing an anchor clears its spill, re-entering a formula there spills afresh and yields to anything typed meanwhile, and undoing the clear restores the array. Saving writes a dynamic array's current rectangle and keeps its `cm` marking, and a legacy Ctrl+Shift+Enter array keeps the rectangle it was entered in, repeating a single row or column and padding with `#N/A`, instead of spilling or shrinking (in Rust, `Sheet::is_dynamic_array` tells the two apart and `evaluate_spill` takes the legacy rectangle as its last argument).
