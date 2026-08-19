---
"@betteroffice/docx": minor
"@betteroffice/docx-react": minor
---

Header/footer editing state becomes one value instead of a flag per kind. `partEdit` names the single non-body part the editor has open, folding the band's first-page variant and the page it was opened from into it — three `useState`s in `DocxEditor` become one.

`PagedEditor` takes `partEdit` in place of `hfEditMode` / `hfEditRId`, and reports the open part's selection through `onYrsPartSelectionChange` in place of `onYrsHfSelectionChange`. `CanvasHfSelectionOverlay` becomes `CanvasPartSelectionOverlay` and queries the display list directly, so `computeHfCaretRectsFromDisplayList` and `computeHfSelectionRectsFromDisplayList` are gone from `@betteroffice/docx/layout`.

Behaviour is unchanged, with one exception: Escape moves from the header/footer chrome component to the paged area, so it now also closes a band whose chrome never mounted.
