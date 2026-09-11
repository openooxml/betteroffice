`paragraph-right-margin.pptx` is a synthetic, one-slide repro for the paragraph
right margin. It contains no private corpus material.

Every body is 600 px wide with zero insets and carries the same sentence, so the
wrap width is the only thing that varies between them.

- `Inherited margins` wraps its first paragraph at 504 px, from the layout
  placeholder's `lstStyle/lvl1pPr` `marR="914400"`, and its second at 552 px:
  the layout defines no level 2, so it falls through to the master's
  `bodyStyle/lvl2pPr` `marR="457200"`.
- `Shape margin` overrides the layout's 96 px from its own `a:lstStyle` with
  `marR="2057400"` and wraps at 384 px.
- `Direct margin` carries the same shape list style. Its first paragraph resets
  the margin with `marR="0"` and wraps at the full 600 px; its second wraps at
  384 px.

`paragraph-right-margin.before.png` and `.after.png` are this slide rendered at
96 DPI without and with the fix.

`marR` is measured from the right edge of the text box, so it narrows the line
without moving its start. LibreOffice breaks all five paragraphs at the same
words. See the [DrawingML paragraph properties](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.paragraphproperties).
