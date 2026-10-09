# betteroffice-ooxml-diff

A bounded longest-common-subsequence diff over token slices, shared by the
BetterOffice formats. It knows nothing about text, styles or documents: callers
tokenize, define token equality and project the result.

`diff_tokens(old, new, limits)` returns ordered hunks of equal, deleted and
inserted token index ranges. Common prefixes and suffixes are trimmed before the
quadratic table is built, the table allocation is checked, and ties prefer
deletion before insertion. `DiffLimits` bounds the token counts and the table
cells: `LimitFallback::Error` refuses an oversized diff, and
`LimitFallback::ReplaceMiddle` reports the untrimmed middle as one deletion and
one insertion instead.

PPTX proposal previews and DOCX comparison use it.
