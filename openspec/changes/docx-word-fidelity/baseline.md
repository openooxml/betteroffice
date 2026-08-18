# Baseline — recorded 2026-08-18, before the instruments landed

## Test suite

`cargo test` (workspace, default features): green, exit 0.

## Fixture inventory (task 0.1)

- Binary OOXML fixtures in `crates/`: **4** — `docx-edit/tests/fixtures/footnote-anchor.docx`, `ooxml-text/tests/fixtures/line-spacing-baseline.docx`, `pptx-parse/tests/fixtures/chart-deck.pptx`, `betteroffice-xlsx/tests/fixtures/hidden-dimensions.xlsx`. Every other DOCX test builds its package from string-literal XML.
- Round-trip assertions before this change: `opens_edits_saves_and_reopens_typed_structure` (typed-structure equality only) and `an_unreadable_chart_part_survives_a_save_byte_for_byte` (one part, byte level), plus per-bug save tests in `crates/betteroffice-docx/tests/document.rs`.
- `docx-layout/tests/goldens.rs`: 14 scenarios, `REQUIRED` = 2; the other 12 may degrade to `UNSUPPORTED` without failing.
- Regeneration gates in use: `DL_SNAPSHOT_UPDATE` (display lists, header/footer bands), `GOLDEN_UPDATE` (raster PNGs), none (hand-edited `.golden.json`, inline chart snapshot).

## Serializer normalizations found (task 0.3)

Established empirically by the first oracle run (unedited open→save of a package the serializer models):

1. Modelled WML parts re-emit Word's standard namespace declaration set (20 URIs) on the root element, regardless of what the input declared. Declared as `root-standard-namespace-declarations`.
2. Modelled WML parts re-emit Word's standard `mc:Ignorable` list on the root element. Declared as `root-mc-ignorable`.

No other structural deviation was observed on the probe fixture. The registry (`ooxml-fidelity::DECLARED_NORMALIZATIONS`) holds exactly these two entries; anything new the corpus surfaces is a bug or a reviewed third entry.

## Defects measured by the first oracle run

1. **Unknown elements inside a modelled part are dropped on save.** A foreign-namespace block element and inline element both vanish; the census reports the exact losses. Pinned by `betteroffice-docx fidelity::an_unknown_element_in_a_modelled_part_is_a_known_defect` (ceiling: 3 lost elements on the probe fixture).
2. **Unknown attributes on known elements are dropped on save.** The census is blind to attributes; the digest's block-attribute record names the loss. Pinned by `fidelity::an_unknown_attribute_on_a_known_element_is_a_known_defect` (ceiling: 1 difference).

Both are recorded on `pkg.unknown-xml` in the ledger. Implementing generic preservation must lower both ceilings to zero.

## Facts the first oracle run established positively

- Non-XML parts (media) and unmodelled XML parts (`customXml/item1.xml`) survive save byte-identically.
- Save → reopen → save is byte-identical on the probe fixture (the serializer is a fixed point on its own output).
- An unedited round trip of fully modelled content passes fingerprint, digest, and census with only the two declared normalizations forgiven.
- `replace_paragraph_text` refuses a paragraph carrying bookmark markers (`UnsupportedParagraphEdit`) — an `edit.typing` data point for the ledger.
