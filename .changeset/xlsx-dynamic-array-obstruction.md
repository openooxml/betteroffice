---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: any value typed, pasted, proposed, batched or written by a raw operation into a dynamic array's spill, the value it already shows included, is now the author's: it is kept and makes the array show `#SPILL!`, as in Excel, instead of being overwritten on the next recalculation, and undo spills the array again. Undoing a retyped anchor formula, a write into a spill or a row or column deletion that cut into an array brings back the result it replaced, including a cached result the engine cannot compute, and redo brings back the newer one. Clearing an anchor clears its spill and its readers recalculate; a formula typed there afterwards is a new entry rather than the old array, and undoing the clear restores the array. Writes into a legacy Ctrl+Shift+Enter array beside its anchor are refused, as Excel refuses them. Saving writes a dynamic array's current rectangle and each anchor's own `cm` index, and a legacy array keeps the rectangle it was entered in, repeating a single row or column and padding with `#N/A`, instead of spilling or shrinking (in Rust, `Sheet::array_definition` tells the two apart and `evaluate_spill` takes the legacy rectangle as its last argument).
