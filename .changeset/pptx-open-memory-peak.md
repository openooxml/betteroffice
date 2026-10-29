---
"@betteroffice/pptx": patch
---

Lower PowerPoint editor memory when opening large decks: temporary copies of the deck state made while opening are released before the next one is built, so the engine's memory peak stays lower for the whole session.
