---
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Typing in the DOCX editor now lands where the click did in a document holding a table, a page or column break, or a block-level content control. Past one of those, a click resolved one position too far to the left per embed ahead of it, so a caret placed mid-word typed into the word before it — and once the drift crossed a paragraph boundary the character landed in another paragraph entirely.

The pointer path crosses two position spaces. The display list gives a block embed a node of its own and opens the next paragraph after it; the engine measures a paragraph from the previous pilcrow, so the same embed counts as one unit *inside* the following paragraph's `Loc` span — which is what `paragraph_spans` reports and what `insertText` resolves against. Both conversions between those spaces dropped those units:

- `YrsPositionProjection` accumulated its input positions in neither space, so its answers drifted one per embed against the map built from the engine's own spans. It now records, per paragraph, the story units standing between the previous pilcrow and its inline content, and carries them across both directions. A block-level column break is also recognised as a block, the way the layout bridge already reads it, rather than as an inline atom.
- The resident caret painted at the paragraph's `pm_start` plus the raw `Loc` offset, which placed it one glyph past the caret for a paragraph sitting directly behind a block embed. The story-index lookup now computes the paragraph-node offset in the same traversal, so the painted caret and insertion point agree without a second full-story scan per frame.
- Resident Backspace and Delete indexed an embed-free text snapshot with an embed-inclusive `Loc` offset. They now inspect the adjacent story segment, deleting a full text code point or one embed and retaining paragraph-boundary merges without an out-of-bounds panic.
