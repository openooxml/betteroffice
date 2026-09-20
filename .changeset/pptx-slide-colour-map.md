---
'@betteroffice/pptx': patch
'@betteroffice/python-pptx': patch
'@betteroffice/rust-crates': patch
---

Resolve `tx1`/`bg1`/`tx2`/`bg2` scheme colours through the slide's colour map: the master's `p:clrMap` and any `p:clrMapOvr/a:overrideClrMapping` on its layout or the slide itself now decide which `a:clrScheme` slot each name reaches, so dark-master decks paint their backgrounds, shape fills and text the way PowerPoint does. One shared resolver feeds the render, snapshot and save projections, so they cannot disagree about a slot.
