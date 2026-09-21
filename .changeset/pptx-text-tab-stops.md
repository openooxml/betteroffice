---
'@betteroffice/pptx': patch
'@betteroffice/rust-crates': patch
---

pptx: a tab in slide text now advances to a tab stop. Text layout had no notion
of tabs at all, so a `U+0009` went to the shaper like any other character and
came back as the fallback face's `.notdef` — a box where that face draws one,
nothing where it does not, and in both cases an advance that owed nothing to the
stops the file declares. Tabbed columns collapsed into the run before them, and
the wrong width moved the wrap point and grew the table row around it.
`a:pPr/@defTabSz` and `a:pPr/a:tabLst` are now read and inherited through the
same slide/layout/master cascade the other paragraph properties use, and a tab
takes the first declared stop past the pen or, failing that, the next multiple
of the default pitch — one inch when nothing declares one — measured from the
text area's left edge, painting no glyph and never reaching past the line. A
hanging indent adds the implicit stop at the paragraph margin that its first tab
lands on; a marker already owns that space here, so only an unmarked hanging
indent has any. 20 of the 103 corpus decks carry tab characters, 17 of them on
slides. On the 103-deck fidelity corpus (101 scored, 1,033 pages) the mean rises
0.89765 to 0.89784, +0.00019; 16 decks improve, led by `pptarena-033` +0.00695,
`pptarena-050` +0.00325, `pptarena-043` +0.00282 and `pptarena-010` +0.00181,
with no deck regressing and page counts exact in both arms.
