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

Open each committed document in Word and export it to PDF. Fill every fixture's
entry in `word.json` with:

- `markerY`: the top of the `MARKER HEADING` line box, in points from page top.
- `anchorLines`: the line count of the first regular body paragraph after the
  table, excluding the caption. For `right-5500-empty-anchor`, count the empty
  anchor's paragraph-mark line, not the later text paragraph.

Record the Word version and export platform here when filling the measurements.
PDFs and measurement logs belong in the PR's review attachments. The baseline
values are intentionally null until the owner measures Word. All five tests
require numeric values; none are ignored or skipped. They compare marker Y
within 2pt and anchor line count exactly.

The test uses the same session and section setup as
`header_footer_reservation.rs`, with the bundled Liberation Sans font for Arial.
The regular face supplies both normal and bold measurement chains; these
assertions concern vertical metrics and line counts, and the short heading
fits beside every table.

## BO prediction from source

These are hand calculations, not Word measurements or results of running BO.
At 11pt, the bundled face's single line pitch is
`11 * (1854 + 434 + 67) / 2048 = 12.64892578125pt`. Empty paragraphs use the
11pt x 1.15 minimum, or 12.65pt. The text contains 61 words. The cleared strips
are 227.7pt for `right-4660` and 225.95pt for `right-4695`; both fit seven words
per line, producing nine lines. Full width fits fifteen words per line,
producing five lines.

`markerY = 72pt + anchorLines * 12.64892578125pt + 12.65pt + 12pt`.
The table's natural height is about 150.898pt, so it still intersects the
heading but leaves enough horizontal room for that single line.

| Fixture | Before the fix | After the fix | Anchor lines before/after |
| --- | --- | --- | --- |
| `right-4660` | 210.490pt | 210.490pt | 9 / 9 |
| `right-4695` | 159.895pt | 210.490pt | 5 / 9 |

The earlier trace's 264pt and 168pt predictions used Courier New 12pt with
exact 16pt lines and fixed row heights. Those values do not apply to these
Arial 11pt fixtures with natural row heights and an empty separator.

## Wrap policy

For a placed table spanning `x` through `x + T` in a column of width `W`,
cleared gaps are `x - leftFromText` and `W - x - T - rightFromText`.
Missing horizontal clearances retain the measurer's 12px default. Tables
wider than half the column use the larger usable gap, with a tie selecting
text on the right. Thus a right-aligned table leaves text on the left and a
left-aligned table leaves text on the right. Narrow tables retain their
existing side and margin arithmetic, including the one-sided choice for
interior tables. Wider interior tables also use one side rather than adding
two-strip flow. That interior policy is an inference awaiting Word measurements.

Measurement and placement share the gap calculation and 24px minimum. If both
gaps are strictly below that minimum, top-level placement clears below the
table and measurement adds no side exclusion, avoiding a second height charge.
Exactly 24px remains usable. Nested table measurement retains its clearance
zone because nested overlays do not use top-level placement's pen advancement.

`line_filler.rs` is unchanged. The empty-anchor fixture is a probe for the
owner's Word export, not evidence for changing empty-paragraph clearance.
