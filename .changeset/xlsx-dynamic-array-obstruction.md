---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: a value typed, pasted, proposed, batched or written by a raw operation into a dynamic array's spill is now kept and makes the array show `#SPILL!`, as in Excel, instead of being overwritten on the next recalculation, and undo spills the array again; clearing the anchor clears its spill. Saving writes a dynamic array's current rectangle and keeps its `cm` marking, and a legacy Ctrl+Shift+Enter array keeps the rectangle it was entered in, repeating a single row or column and padding with `#N/A`, instead of spilling or shrinking (in Rust, `Sheet::is_dynamic_array` tells the two apart and `evaluate_spill` takes the legacy rectangle as its last argument).
