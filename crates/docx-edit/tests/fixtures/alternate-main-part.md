# Alternate main-document filename

This original fixture contains the same visible content as `anchored-header.docx`, with its main part named `word/document2.xml`. The package's `officeDocument` relationship and content-type override point to that part; its relationships live in `word/_rels/document2.xml.rels`.

Before the fix, the parser silently returned an empty body and rendered a blank page. The regression test checks both parsing facades, header relationships, full saves, and selective paragraph saves. Edits must replace `document2.xml`, without adding an unreferenced `document.xml` or a mismatched relationships part.

Tests generate the fixture in memory through `tests/support/quality_fixture.rs`. Run `cargo test -p betteroffice-docx-edit --test alternate_main_part`; no document files are committed or written by the test.

A relocated variant uses `custom/document2.xml` with relative and absolute header targets under `stories/`; edits survive full and selective saves without creating stray `word/` parts.
