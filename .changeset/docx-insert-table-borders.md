---
"@betteroffice/rust-crates": patch
"@betteroffice/docx": patch
---

Inserting a DOCX table now produces one with borders. `insert_table` gave each new cell only its spans, so a table dropped into a document whose styles carry no bordered table style rendered as bare text with no rules at all — and saved that way too, since the writer lifts `w:tblBorders` out of the cells it finds.

Every cell of a new table now carries Word's Table Grid rules: a single 0.5pt (`w:sz="4"`) black line on each of its four edges, so the grid reads the same as the one Word draws for its own inserted table. Resolved table borders have always lived on the cells in this engine — seeding pushes `w:tblBorders` down per position, the display list paints only what `tcPr.borders` names, and the toolbar reads its border state from there — so authoring them at the cell is what makes the table visible, survive a save, and come back bordered when the file is reopened.

Referencing a style instead would not have worked: nothing synthesises a `TableGrid` definition into `styles.xml`, so a blank document has none to inherit from. Rows and columns added later still inherit the borders, since both copy their template cell's `tcPr`, and a suggested insertion carries them exactly like a direct one.
