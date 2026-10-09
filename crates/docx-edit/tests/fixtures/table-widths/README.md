Run `python3 generate.py` to rebuild these synthetic DOCX fixtures. They use Arial 12 pt with exact 12 pt line spacing, zero paragraph spacing, A4 pages, and 80 table rows.

`word.json` records Microsoft Word 16.113.2's PDF page boundaries and cell line endings. The PDFs were exported with `word_export.py` and extracted with `word_lines.py`. Every fixture has three pages starting at rows 1, 30, and 59. The expected cell text repeats with the row number substituted.

The fixtures cover first-row preferred widths, a spanning first-row cell, explicit autofit, omitted layout, a cell with `noWrap`, and cell margins with table and paragraph indentation.
