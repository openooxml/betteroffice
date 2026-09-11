---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Skip hidden slide shapes and hidden groups' descendants when painting and hit-testing.

Deck schema 6 migrates existing version 1–5 documents by recovering hidden flags from stored package data after the schema-5 theme-formatting migration. Older clients reject the new schema.

`ShapeSnapshot.hidden` is optional and omitted when false, preserving unchanged snapshot JSON. Only hidden shapes store a Yrs key; the schema stamp changes for all decks.
