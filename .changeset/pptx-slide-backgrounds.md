---
"@betteroffice/pptx": minor
"@betteroffice/rust-crates": minor
---

PPTX slides can declare their own background: solid or gradient `ShapeFill` values write `p:bg` through `SlideWrite`, `DeckSession::set_slide_background` and the typed `setSlideBackground` batch step, and the editor paints the change before the deck is saved. A `null` background removes the slide's declaration, including fills the model cannot represent. Invalid fills are refused before they reach the save, and an invalid fill from a peer is ignored in favour of the file's own background.
