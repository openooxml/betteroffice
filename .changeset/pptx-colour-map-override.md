---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Honour the presentation colour mapping. A master's `p:clrMap` and a layout's or slide's `p:clrMapOvr/a:overrideClrMapping` now decide which theme slot `bg1`, `tx1`, `bg2`, `tx2` and the accents stand for, so a dark layout — the one that maps `bg1` onto `dk1` and `tx1` onto `lt1` — paints white text on a black slide, as PowerPoint does, rather than the light inverse of it. A slide carrying `a:masterClrMapping` inherits the mapping from its layout, and a layout carrying it inherits from its master.
