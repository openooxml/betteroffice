# Floating table wrapping

These synthetic documents isolate the half-column cutoff for floating tables.
Regenerate them with `python3 generate.py` from this directory. The generator
uses Python's standard library, stores ZIP entries with fixed timestamps and
permissions, and embeds the same 1x1 PNG in every document. It leaves `word.json`
untouched.

Each document has a 11907 x 16840 twip page with margins top 1440, right 1134,
bottom 1134, left 1418, header 709 and footer 709. The text column is 9355 twips
(467.75pt). Minimal document defaults select Arial 11pt, single line spacing,
and zero paragraph spacing. No paragraph names a style.

Each floating table is text anchored with `tblpY="1"`, 141 twips of left and
right text clearance, `tblOverlap="never"`, automatic preferred table width,
a fixed one-column grid, and 28 twip cell margins on every edge. Rows have
natural heights. The first row contains one inline picture whose `wp:extent`
is 120pt x 120pt. The second contains a caption separated by `w:br`.

| Fixture | Alignment | Grid and cell width | First body paragraph after the table |
| --- | --- | --- | --- |
| `right-4660.docx` | right | 4660 twips | 350 ASCII characters |
| `right-4695.docx` | right | 4695 twips | 350 ASCII characters |
| `right-5500.docx` | right | 5500 twips | 350 ASCII characters |
| `left-5500.docx` | left | 5500 twips | 350 ASCII characters |
| `right-5500-empty-anchor.docx` | right | 5500 twips | empty, then the 350-character paragraph |

The body text is 15 repetitions of `alpha beta gamma delta ` followed by
`alpha`, exactly 350 characters. An empty separator follows it. The last
paragraph contains exactly `MARKER HEADING`, bold 12pt with 240 twips (12pt)
of space before.

## Word measurements

`word.json` holds, per fixture, the `MARKER HEADING` baseline in points from
the page top and the line count of the first body paragraph after the table
(for `right-5500-empty-anchor`, the empty paragraph's single line). Values come
from PDFs exported by Word for Mac 16.113. Word wraps text beside the table on
the side away from it at every width, so `right-4660` and `right-4695` match,
and it lays the empty anchor paragraph out beside the table rather than below
it. The test compares the baseline within 2pt and the line count exactly.
