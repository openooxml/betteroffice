---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
"@betteroffice/python-pptx": patch
---

Render slide text closer to PowerPoint: autofit keeps point spacing and rounds sizes to whole points, automatic numbering skips blank paragraphs, text is centred within its line spacing, named families use their own text metrics and heavy weights, text outside placeholders takes the presentation's `defaultTextStyle`, a character a face lacks is drawn from a script fallback face, and right-to-left paragraphs are laid out from the right.
