---
'@betteroffice/pptx': patch
'@betteroffice/rust-crates': patch
---

pptx: resolve PowerPoint's built-in table styles, and stop dressing a table in
the one `tableStyles.xml` only nominates. `ppt/tableStyles.xml` carries the
styles a deck edited, so a table naming a style PowerPoint has never had to
write out — `{5C22544A…}` "Medium Style 2 - Accent 1" above all, the default
every new table and every python-pptx table takes — found nothing and rendered
as bare text on the slide background: no header fill, no banding, no borders,
no white bold header. 17 of the 59 tables in the 103-deck corpus, across 11
decks, name a style their own package never defines. Three of those styles are
now resolved from a built-in catalogue, transcribed from the definitions
PowerPoint itself serialises when a deck does edit them. Separately, the
`def` attribute names the style the authoring UI hands a *new* table, not a
fallback for one that names none, so a table with an empty `a:tblPr` no longer
picks it up: `pptarena-010` was painting ten converted forms in `accent1`
tint 20% where PowerPoint leaves them white.

On the 103-deck fidelity corpus (101 scored, 1,033 pages) the mean rises
0.89784 to 0.89900, +0.00116; 9 decks improve, led by `pptarena-061` +0.04292,
`pptarena-071` +0.03352 and `pptarena-025` +0.02781, with no deck regressing,
92 of 101 decks byte-identical, and page counts exact in both arms.
