# layout-fidelity

## ADDED Requirements

### Requirement: The layout golden gate fails hard and only grows

The `goldens.rs` required list SHALL be a ratchet: a scenario, once required, stays required, and the required count lives in the ledger so shrinking it fails the ledger's exact-equality check. A scenario the builder cannot handle MAY report as a gap only while it is not yet required; the ledger records the gap count, and that count may only fall.

#### Scenario: A supported scenario cannot regress to unsupported

- **WHEN** a required scenario starts erroring instead of laying out
- **THEN** the suite fails; degrading to a gap is not available to required scenarios

#### Scenario: Coverage growth is one-way

- **WHEN** a previously unsupported scenario first passes and is added to the required list
- **THEN** the ledger's required count rises, and any later removal fails the ledger check

### Requirement: Corpus fixtures pin the geometry that decides them

Every corpus fixture with a layout oracle SHALL pin page count and drawing count, and — for the scenario the fixture exists for — the specific geometry that decides it: line boxes for a wrap fixture, row split positions for a table fixture, band geometry for a header fixture. A page count alone is a weak gate: it can hold while every line breaks in the wrong place.

#### Scenario: The deciding geometry is the assertion

- **WHEN** a float-wrap fixture lays out
- **THEN** the test asserts the wrapped lines' boxes, not only that the document still makes N pages

### Requirement: Positioning follows Word's rules

Anchored objects SHALL be positioned per their declared bases — page, margin, column, character — with Word's resolution of alignment, offsets, and overlap; floating tables and text wrap SHALL push text the way Word pushes it; and the layout criteria in the catalogue (line breaking, spacing rules, keeps, column balancing, footnote space, header growth) SHALL each carry at least one pinned fixture before their layout axis claims `pinned`.

#### Scenario: An anchored image lands where Word lands it

- **WHEN** a corpus fixture anchors an image relative to the margin with an offset
- **THEN** the display list places it at the position recorded in the fixture's pinned oracle, established against Word's rendering of the same file

### Requirement: Hand-written vectors are never regenerated from engine output

Golden vectors established by review — against ECMA-376, font metrics, or Word's rendering — SHALL only change through the same review, never by regenerating from the engine. Regenerating a golden from the code under test turns the oracle into a mirror.

#### Scenario: A layout change meets its reviewer

- **WHEN** a layout change moves a hand-written vector's expected value
- **THEN** the vector is edited by hand, with the change reviewed against the reference, not regenerated

### Requirement: One regeneration gate

Every regenerable golden in the workspace SHALL regenerate only under `GOLDEN_UPDATE=1`. `DL_SNAPSHOT_UPDATE` is retired; hand-maintained goldens that are in fact regenerable (the layout `.golden.json` files, the inline chart snapshot) move under the same gate. Drift stays a deliberate act with one name.

#### Scenario: No golden updates by accident

- **WHEN** the suite runs without `GOLDEN_UPDATE=1` and output differs from a golden
- **THEN** the test fails and no golden file is written

### Requirement: Layout is deterministic

The same document laid out twice SHALL produce identical output, byte-for-byte at the display-list serialization. Nondeterministic layout makes every golden flaky and every diff unreadable.

#### Scenario: Two runs, one answer

- **WHEN** any corpus fixture is laid out twice in one process
- **THEN** the serialized display lists are identical
