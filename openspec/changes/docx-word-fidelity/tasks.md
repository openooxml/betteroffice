# Tasks

## Implementation evidence

Phases 0–1 landed with the first oracle run: the crate, its unit tests, and the round-trip gates on a synthetic probe fixture. The run itself produced the baseline's findings — two declared serializer normalizations, byte-stable fixed point, and the unknown-markup drop pinned as the first ledger defects. The corpus (phase 4) is what turns the phase-2 gates from probe-fixture claims into corpus claims.

## 0. Baseline before code

- [x] 0.1 Record the current state as numbers: binary OOXML fixtures in `crates/` (4), round-trip assertions and what they compare, `goldens.rs` required count (2 of 14), regeneration gates in use (`DL_SNAPSHOT_UPDATE`, `GOLDEN_UPDATE`, none)
- [x] 0.2 Run `cargo test` and record the green baseline the phases are measured against
- [x] 0.3 Enumerate every intentional serializer normalization currently known, as the seed of the D5 registry
- [x] 0.4 Seed `ledger.json` honestly: every criterion in the catalogue, statuses as they are today, no gate names invented

## 1. Instruments — `crates/ooxml-fidelity`

- [x] 1.1 Create the crate: neutral XML tree reader over part bytes (namespace-aware, bounded depth and size, no external fetches), independent of `docx-parse`
- [x] 1.2 Structural fingerprint per `specs/fidelity-oracles`: URI-resolved names, sorted attributes, `xml:space`-aware whitespace significance, node order preserved
- [x] 1.3 Normalizations registry (D5): the declared list, applied by the fingerprint, reviewed entry by entry
- [x] 1.4 Element census by local name across every XML part; `losses(before, after)` reports only counts that shrank
- [x] 1.5 WML semantic digest in a `wml` module: story discovery, block records (path, attributes, text, nested property tokens, generic-subtree fingerprints), structure walk over storyless parts
- [x] 1.6 `diff_digests` returning `{path, before, after}` entries, ordered
- [x] 1.7 Comparison-mode registry (D6): `exact`, `canonical-exact`, `tolerance` with declared epsilon; unknown artifact refuses
- [x] 1.8 Unit tests per mechanism, including the digest special cases: hyperlinks, bookmarks, drawings, content controls, foreign-namespace property children

## 2. Round-trip harness — `crates/betteroffice-docx/tests/fidelity.rs`

- [x] 2.1 Round-trip helper: open, save, reopen the produced bytes, both sides parsed by `ooxml-fidelity`
- [ ] 2.2 Unedited gate: fingerprint equality, empty digest diff, zero census loss, over every existing fixture and the demo document (landed on the probe fixture; existing fixtures and the demo document join with the corpus)
- [x] 2.3 Byte rules: non-XML parts and unmodelled XML parts byte-identical through save
- [x] 2.4 Fixed point: save → reopen → save byte-identical
- [ ] 2.5 Edited gate: the canonical edit set (type a character, split a paragraph, set a paragraph property, set run properties over a range) leaves the census loss-free and the digest diff explainable by the edit (landed for the text edit; the remaining three operations follow)

## 3. Acceptance and comprehensive fixtures

- [ ] 3.1 Export the supported run-property and paragraph-property inventories from `docx-parse` as data
- [ ] 3.2 `scripts/create-acceptance-fixture.ts`: every inventory entry, authored whitespace, tabs and breaks, unknown markup at paragraph level / in a run / nested in known containers, one page-crossing paragraph; pinned zip date; regeneration command in the header
- [ ] 3.3 Acceptance test: coverage derived from the inventories, checked-in bytes equal the builder output, unedited round trip passes both oracles
- [x] 3.4 Generate the comprehensive document with numbered sections (`scripts/create-wordprocessingml-comprehensive.ts`); Word-authored provenance upgrades follow once the file is validated in Word (open without repair, section screenshots)
- [ ] 3.5 Comprehensive oracles: pinned block counts and per-section assertions that cite section numbers

## 4. Corpus — `crates/betteroffice-docx/tests/corpus/`

- [x] 4.1 `manifest.json` schema and manifest test: SHA-256 per fixture, provenance, features, bound oracle, evidence status; a fixture without an oracle fails
- [x] 4.2 Corpus runner: unedited gate and fixed point over every manifest entry, findings pinned exactly per fixture under `GOLDEN_UPDATE=1`; the edited gate joins with the canonical edit set
- [ ] 4.3 Wire the redaction intake: document in `AGENTS.md` that an engine PR's repro file joins the corpus, through `betteroffice-redact` when confidential
- [x] 4.4 Seed the corpus: the comprehensive and demo documents; real-world redacted documents and the acceptance fixture follow
- [x] 4.5 Builder-drift gate: every generated fixture equals its builder's output, run where the builders run (`scripts/corpus-fixtures.test.ts`)

## 5. Meta-oracles

- [ ] 5.1 Reach: per corpus fixture, digest block count equals the part's block count
- [x] 5.2 Blind-spot pairs: numbering identity, emptied borders, tab stops, paragraph-mark properties, table flattened to paragraphs, section setup, an emptied content control, `styles.xml` and `numbering.xml` definitions — each pair must digest differently
- [x] 5.3 Teeth: strip an element from a saved package; census and digest both fire
- [x] 5.4 Floor: every corpus sweep asserts the fixture count exceeds the pinned floor
- [x] 5.5 Ceilings: known defects probed and pinned with `==` in `defects.rs`, the ledger's enumeration equal to the set that still reproduces, each criterion naming its ceiling test

## 6. Editing fidelity

- [ ] 6.1 Reachability gate: over the corpus, every story a caret can enter in Word — body, nested tables, text boxes, headers and footers, footnotes and endnotes, comments — accepts a caret, and hit testing, caret geometry, and painted glyph positions agree at every block's first and last position
- [ ] 6.2 Edit-scope fingerprint in `ooxml-fidelity`: fingerprint with the edited subtree replaced by a sentinel; gate every `docx-edit` operation family with it
- [ ] 6.3 Structural-key parity tests: Enter, Tab, Shift+Tab, Backspace, Delete in plain paragraphs, list items (continue, change level, clear numbering), table cells (navigate, never destroy structure), and across paragraph and table boundaries — each asserting the resulting saved package, not just the model
- [ ] 6.4 Undo gate: for every operation family, undo then save passes both oracles against the pre-edit save
- [ ] 6.5 Tracked-reject gate: edit under tracking, reject all, digest (adjacent identical runs collapsed) equals the original
- [ ] 6.6 Convergence gate: the existing `docx-edit` corpora converge to a state whose save passes the fidelity gates
- [ ] 6.7 Lock and placeholder semantics: content-control locks refuse the operations Word refuses; placeholder text is replaced wholesale on first input, never appended to

## 7. Scorecard — `crates/betteroffice-docx/tests/scorecard/`

- [x] 7.1 `ledger.json` implementing the catalogue: one entry per criterion, six axes, per-axis target, gates, defects
- [x] 7.2 Ledger test: id set equals the catalogue, no status outranks its target, measured numbers recomputed and compared exactly, every claimed status's gate and every defect's ceiling exists, regeneration only under `GOLDEN_UPDATE=1` and refusing a number that moved the wrong way
- [ ] 7.3 Generated `scorecard.md`: per-axis criteria-at-target counts, corpus metrics, defect list; drift-checked against the ledger
- [ ] 7.4 Retire `DL_SNAPSHOT_UPDATE`; gate the hand-edited `.golden.json` files and the inline chart snapshot under `GOLDEN_UPDATE=1`

## 8. Layout and visual

- [ ] 8.1 `goldens.rs`: move the required list's count into the ledger; additions only
- [ ] 8.2 Per-corpus-fixture layout pins: page count, drawing count, and the geometry that decides each scenario
- [ ] 8.3 Expose the bytes→measure→layout lowering so `.docx` bytes reach `Layout` and `DisplayList` without a caller-supplied projection
- [ ] 8.4 Corpus render goldens on `docx-raster`: exact byte compare; failure writes actual, expected, and diff images
- [ ] 8.5 Register the raster tolerance path (if any survives) in the D6 registry with its epsilon

## 9. Word evidence

- [ ] 9.1 Write the probe protocol into the corpus README: probe → Word round trip → paired fixture → manifest record of what Word preserved, normalized, refused
- [ ] 9.2 First paired fixtures for the behaviors the engine currently assumes without evidence; update the ledger's evidence axis accordingly
