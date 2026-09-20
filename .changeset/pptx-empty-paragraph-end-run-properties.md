---
'@betteroffice/pptx': patch
'@betteroffice/rust-crates': patch
---

pptx: an empty paragraph now takes the height its own `a:endParaRPr` asks for. Empty paragraphs are how authors write vertical spacers, and PowerPoint sizes each one from the run properties it carries; the renderer read the size the shape's list style would have given a run instead, so on `pptarena-042` a 40pt spacer was drawn at the 211.2pt title default — 5.28 times too tall — and the error accumulated down the body until text left its placeholder and crossed the footer. The property was parsed and never read. It is now resolved through the same slide/layout/master cascade the rest of the paragraph properties already use, so it reaches text that arrives as a collaborative story as well as text read straight from the package, and a paragraph without one keeps the size it inherits today. On the 103-deck fidelity corpus (101 scored, 1,033 pages) the mean rises 0.89198 to 0.89765, +0.00568; 42 decks improve, led by `pptarena-052` +0.1101, `pptarena-055` +0.0736 and `pptarena-042` +0.0625, and one deck moves -0.000042 while its text lands 4px closer to PowerPoint.
