# Word compatibility spacing

`compatibility-spacing/generate.py` creates seven synthetic DOCX fixtures. The
paired fixtures differ only in their `w:compat` flag. Their Arial 12pt text uses
exact 12pt line spacing on US Letter pages with 72pt margins. Compatibility mode
is 14 so Word preserves spacing after authored page breaks.

`compatibility-spacing/word.json` contains page indices and text baselines from
Word 16.113.2 PDF exports, measured in points from the page top with
`word_lines.py`.
The PDF exports remain in `/tmp/sol-compatspacing-word`, outside git.
`tests/compatibility_spacing.rs` compares relative advances and page starts,
avoiding differences in font baseline metrics.

| Fixture pair | Word default | Word with flag |
| --- | --- | --- |
| `auto` | 14pt automatic gaps before and after | 5pt before, 10pt after |
| `break` | 24pt before the leading hard-break target | 0pt before that target; `pageBreakBefore` keeps 24pt |
| `table` | 12pt advances between same-style cell paragraphs | 30pt advances, including 18pt paragraph spacing |

The auto fixtures supply 100/200-twip fallback spacing as well as auto-spacing
attributes. Word uses those fallback values with
`doNotUseHTMLParagraphAutoSpacing`. The additional `auto-without-fallback`
fixture records zero gaps in Word when the flag is set and fallback spacing is
absent. BetterOffice implements the requested fixed 5pt/10pt interpretation
from [Microsoft's element reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.donotusehtmlparagraphautospacing?view=openxml-3.0.1);
the no-fallback test asserts that interpretation, not Word equivalence.
BetterOffice's unflagged 14px constant is preserved; Word's 14pt default is
recorded separately.

`suppressSpBfAfterPgBrk` suppresses spacing on paragraphs whose first run is a
hard page break and which have subsequent content. `pageBreakBefore` and an
independent break paragraph retain their existing behavior. This matches the
[element definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.suppressspacingbeforeafterpagebreak?view=openxml-3.0.1).
A separate Word probe also suppressed spacing after a leading column-break run
in a single-column section. That behavior is not implemented here: the seed
currently records leading column breaks as separate blocks without the
`pageBreakBeforeRun` marker, so handling it requires coordinating the break
representation with the pagination work.

`allowSpaceOfSameStyleInTable` disables `contextualSpacing` in cells, including
in nested tables, while body paragraphs keep contextual suppression. It adds
spacing when enabled, as described by
[Microsoft's reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.allowspaceofsamestyleintable?view=openxml-3.0.1).
Unflagged table measurement retains the pre-existing behavior: contextual
suppression runs during placement, after row extents are measured. For this
fixture, both variants therefore measure a 96px row and place `TABLE END` at
240px. The default paints cell paragraphs with 16px (12pt) advances inside that
row; enabling the flag paints them with 40px (30pt) advances. The test locks the
default measurement to the parent revision and compares cell advances in both
variants, plus the enabled body-to-`TABLE END` advance, against Word. Closing the
gap between unflagged row measurement and contextual cell placement is outside
this compatibility change.

The table compatibility survey also found missing grid and wrapping settings:

- [`adjustLineHeightInTable`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.adjustlineheightintable?view=openxml-3.0.1)
  applies section line-grid pitch to cell text; it is normally excluded.
- [`doNotSnapToGridInCell`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.donotsnaptogridincell?view=openxml-3.0.1)
  disables grid snapping in cells containing floating objects.
- [`layoutRawTableWidth`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.layoutrawtablewidth?view=openxml-3.0.1)
  excludes table indentation from the fit calculation beside a floating object.

These are not parsed or implemented by this change. The grid settings need
consistent per-section propagation through measurement and cell layout;
lowering alone cannot implement them because later grid-resolution passes
overwrite the pitch. The wrapping setting belongs to placement in the
pagination-owned area. `useWord2002TableStyleRules` affects conditional table
border placement rather than paragraph spacing. Table row wrapping and
breaking flags likewise belong to the pagination work.
