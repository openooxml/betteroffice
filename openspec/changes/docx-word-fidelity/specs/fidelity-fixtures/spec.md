# fidelity-fixtures

## ADDED Requirements

### Requirement: The acceptance fixture covers the supported-property inventory and proves it

A generated acceptance document SHALL cover, deliberately and in one file: every supported run property and every supported paragraph property, each in its own labelled run or paragraph; authored whitespace that must survive verbatim; tabs and hard breaks; unknown markup interleaved with known content at paragraph level, inside a run, and nested inside known property containers; and enough text that a single paragraph crosses a page boundary. `docx-parse` SHALL export its supported-property inventories as data, and the acceptance test SHALL derive its expected coverage from them.

#### Scenario: A new property cannot ship uncovered

- **WHEN** a run property is added to the supported inventory without regenerating the acceptance fixture
- **THEN** the acceptance test fails on the derived-coverage check

#### Scenario: Unknown markup survives in every position

- **WHEN** the acceptance fixture round-trips unedited
- **THEN** every interleaved unknown subtree fingerprints identically before and after, in position

### Requirement: The comprehensive document is Word-authored and section-numbered

One comprehensive document SHALL be authored in Microsoft Word, with numbered sections, exercising at minimum: the style hierarchy, character formatting, bulleted and numbered lists including restarts, tables with vertical merges, grid spans, and border conflicts, inline and anchored images across wrap modes, text boxes, headers and footers including even/odd and first-page variants, footnotes and endnotes, a table of contents and cross-reference fields, hyperlinks and bookmarks, tracked changes, comments with replies, content controls, and multiple sections with different page setups and column counts. Tests SHALL cite the section numbers they assert against, so the fixture is legible as a specification of itself.

#### Scenario: A layout test cites its section

- **WHEN** a test asserts the resolved border of the comprehensive fixture's conflict table
- **THEN** the test names the section number that table lives in

#### Scenario: The comprehensive fixture is fully gated

- **WHEN** the corpus runner executes
- **THEN** the comprehensive document passes the unedited and edited gates, and its layout pins (page count, drawing count) hold

### Requirement: Word evidence comes from paired probe fixtures

When ECMA-376 does not decide a behavior, the behavior SHALL be established by experiment: author a probe document, round-trip it through Microsoft Word, and check both files in as a pair — `<name>-probe.docx` and `<name>-word-roundtrip.docx`. A test SHALL assert the engine reads both members identically for the probed behavior, and the manifest entry SHALL record what Word preserved, normalized, or refused, with the Word variant and date of the experiment.

#### Scenario: An assumption is upgraded to evidence

- **WHEN** the engine's handling of a construct rests on an undocumented assumption about Word
- **THEN** a probe pair is added, the assumption is confirmed or corrected, and the ledger's evidence axis moves from `spec` or `none` to `word`

#### Scenario: Word's own normalization is not a false alarm

- **WHEN** Word's round trip of the probe applies a normalization (recorded in the manifest)
- **THEN** the pair test asserts recognition equality on the probed behavior rather than byte equality of the pair
