`slide-number-fields.pptx` derives from the tracked BetterOffice demo deck.

- `ppt/presentation.xml` starts numbering at 10.
- `ppt/slides/slide1.xml` has `show="0"`; its position still counts.
- Master shape 50 carries `slidenum` with cached `MASTER-CACHE`, followed by
  ordinary `|` text and a `datetime` field cached as `CACHED-DATE`.
- Layout shape 51 carries `slidenum` with cached `LAYOUT-CACHE`.
- Slide 2 shape 52 carries `slidenum` with cached `77`, deliberately different
  from its effective number, 11. Its editable story retains the cache.

The inherited fields render as 10, 11, and 12. The master number's run specifies
Arial 23 pt, bold, italic, underline, and `#FF0066`; the layout number's run
specifies Arial 17 pt and `#1122CC`. The date's run specifies Arial 13 pt and
`#00AA55`.
