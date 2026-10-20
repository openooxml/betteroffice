---
"@betteroffice/pptx": minor
"@betteroffice/rust-crates": minor
---

PPTX slides can set or clear their own solid or gradient background through `DeckSession::set_slide_background` and the `setSlideBackground` batch step. Edits paint before saving, clearing removes any `p:bg`, and invalid fills, including a peer's, are refused or ignored.
