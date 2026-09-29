---
"@betteroffice/rust-crates": patch
---

Shift the positions of the pages after an edit without cloning a key string per primitive. After a keystroke every display page past the edited one has its body positions shifted, and looking each primitive's block up in the shift table cloned its block key first: an allocation per primitive, ~270 thousand per keystroke on a 267-page document. The first keystroke after opening paid most for it (~210 ms in wasm, where the allocator is slow on the heap the open leaves); later keystrokes paid ~20 ms.
