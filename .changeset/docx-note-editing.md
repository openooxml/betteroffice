---
"@betteroffice/docx": minor
"@betteroffice/docx-react": minor
---

Clicking a footnote or an endnote now opens it for editing. The caret lands under the pointer, typing goes into that note, and what is typed reaches the saved file. Escape leaves the note; clicking back in the body leaves it and places the body caret under the pointer.

Undo follows the part being edited: the first edit in a newly opened part replaces the undo scope and discards the previous part's history.

A note area is document text, not chrome, so a single click opens it — where a header or footer band, which repeats on every page and sits in the margin, still needs a double-click. The click that opens the note also places its caret: the position it resolved belongs to a story the editor is not on yet, so it waits for that story to become active and is discarded if the host never opens the note.

The editor could previously be in a header/footer band, and now also in a note, so the two are one value rather than a flag each: `partEdit` names the single non-body part that is open, and opening a note closes whatever was open before. That removes the state where a band and a note could both claim the caret, and folds the band's first-page variant and originating page into the same value. `PagedEditor` takes `partEdit` in place of `hfEditMode` / `hfEditRId` and reports the open part's selection through `onYrsPartSelectionChange`.

Everything the body offers behind an open band was already off, and stays off behind an open note for the same reason plus one of its own: the display list indexes images, tables and hyperlinks by band, so a note area has no scope to look one up in and answering `body` would hand back something from behind the note. Selecting a picture, dragging a table edge, the row/column insert affordance and opening a link are therefore inert inside a note; caret movement, word and paragraph selection, drag-selection and Home/End are not. Arrows navigate a note the way they navigate a band — paragraph by paragraph, since the display-list vertical-move query answers body positions only — and the caret never leaves the note it is in.

Selection geometry comes from the note's own story: `noteCaretRects` is the note twin of `hfCaretRects`, resolving the leading edge of the caret's position and falling back to the trailing edge of the previous one at end of line. A note paints on exactly one page, so unlike a band there is never a second candidate page to choose between. One overlay now paints the caret and highlight for whichever part is open, band or note, and it asks the queries directly — `computeHfCaretRectsFromDisplayList` and `computeHfSelectionRectsFromDisplayList`, which had that overlay as their only caller, are gone from `@betteroffice/docx/layout`.
