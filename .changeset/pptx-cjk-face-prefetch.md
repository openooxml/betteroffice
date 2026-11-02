---
'@betteroffice/pptx': patch
---

The PowerPoint editor starts downloading the CJK coverage font the first slide needs while its WebAssembly engine is still loading, instead of after the deck is parsed, so Japanese, Korean and Chinese decks show sooner on slow connections. The first frame is unchanged.
