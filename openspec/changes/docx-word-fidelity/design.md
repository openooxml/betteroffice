# Design

## Definitions

- **Fidelity** — three prongs. *Preservation*: a document opened, edited, and saved loses nothing it did not intend to change. *Rendering*: what the engine models is positioned, laid out, and painted the way Microsoft Word does. *Editing*: everything Word lets a user edit, the engine lets a user edit, with the same observable result — same caret behavior, same structural outcome, same file. 100% fidelity is all three, over the whole corpus, with zero known defects.
- **Criterion** — one WordprocessingML feature lane in the catalogue (`specs/fidelity-scorecard/spec.md`), scored on six axes: preserve, model, edit, layout, paint, evidence.
- **Oracle** — a mechanism that turns a fidelity failure into a test failure with a location. An assertion that cannot name what was lost is not an oracle.
- **Reference** — ECMA-376 Part 1 where it speaks; Word's observed behavior where it is silent. The evidence axis records which of the two backs each claim.

## Decisions

### D1 — Fidelity is measured on bytes, not on the model

The oracles parse the *saved package bytes* with a neutral XML reader in `ooxml-fidelity`, and compare against the *original package bytes* parsed the same way. They never read `docx-parse`'s typed model.

Why: an oracle built on the code under test inherits its blind spots. If the typed model drops an attribute, a model-to-model comparison agrees with itself and reports nothing. Parsing both sides from bytes makes the oracle blind only to what XML itself cannot express.

Alternative rejected: comparing typed structures (the current `structure()` equality). It is a useful smoke test and stays, but it is not a fidelity oracle — its reach is exactly the model's reach.

### D2 — Two oracles, and one may not compensate for the other

The **structural fingerprint** answers "is this the same tree?": a projection of each XML part that resolves namespace prefixes to URIs, sorts attributes, resolves QName-valued attributes through in-scope bindings, drops insignificant inter-element whitespace under `xml:space` semantics, and keeps element order, text, and namespace bindings significant.

The **semantic digest** answers "did the round trip keep the meaning?": per story part, an ordered list of block records — containment path, paragraph text with tabs and hard breaks as characters, property tokens that include their nested children, and whole-subtree fingerprints for unknown markup — plus a structure walk that also covers parts with no story (`styles.xml`, `numbering.xml`, `settings.xml`).

Both must pass. A change may not loosen one because the other is green: the fingerprint cannot say a self-consistent tree lost meaning, and the digest deliberately tolerates tree residue (a split run saying the same text) that the fingerprint would reject. Each covers the other's tolerance.

Alternative rejected: byte equality of modelled XML parts. It rejects harmless lexical normalization while saying nothing precise about semantic loss. Byte equality is reserved for what D4 assigns it to.

### D3 — The digest is compared across a real reopen, and diffs are paths

The digest of the opened package is compared to the digest of the package obtained by *reopening the produced bytes*. Digesting an in-memory value against itself is self-consistent and catches nothing. `diff` returns a list of `{path, before, after}` — "the round trip lost something" is useless without saying what, and the failure this oracle exists for is a silent drop; a bare boolean reproduces the silence.

### D4 — Byte rules

1. Non-XML parts — media, embedded fonts, VBA, OLE payloads — SHALL be byte-identical through save.
2. XML parts the engine does not model SHALL be byte-identical through save.
3. Modelled XML parts re-emit normalized, gated by both D2 oracles.
4. Save → reopen → save SHALL be byte-identical (the serializer is a fixed point on its own output).

Rule 4 is what makes "open and save without editing" a safe operation for a user to repeat.

### D5 — Intentional normalizations are declared, or they are bugs

Any deviation the serializer intentionally makes from input XML — attribute-order canonicalization, prefix choice, empty-element spelling — SHALL be listed in a normalizations registry in `ooxml-fidelity`, and the fingerprint applies exactly that list before comparing. An undeclared difference is a failure. The digest never normalizes meaning: there is no such thing as an intentional semantic change on save.

### D6 — Comparison modes are frozen per artifact

A registry pins how every fidelity artifact is compared: `exact` (every field), `canonical-exact` (after declared canonicalization), or `tolerance` (numeric epsilon, allowed only for raster comparison and only with a declared epsilon). Requesting a comparison for an artifact outside the registry is an error. Nobody quietly loosens a comparison to make a test pass.

### D7 — The corpus is manifest-pinned and every fixture has an oracle

The corpus manifest records, per fixture: SHA-256, provenance (`word` | `libreoffice` | `onlyoffice` | `generator` | `redacted`), producing application and version when known, the features it exercises, the oracle bound to it, and its evidence status. A test verifies every hash and refuses a fixture with no oracle. Counts in the manifest are derived, never hand-written. Generated fixtures come from deterministic builders — pinned zip timestamps, regeneration command in the builder header — and a test asserts checked-in bytes equal the builder's output.

Why the binding rule: a fixture nobody asserts anything about documents nothing and rots silently.

### D8 — The acceptance fixture self-validates against the supported-property inventory

`docx-parse` exports its supported run-property and paragraph-property inventories as data. The generated acceptance fixture covers every entry, plus authored whitespace that must survive verbatim, tabs and hard breaks, unknown markup interleaved at paragraph level, inside a run, and nested inside known property containers, and enough text that one paragraph crosses a page. Its test derives expected coverage from the inventories — supporting a new property without regenerating the fixture fails the build.

### D9 — Meta-oracles: the oracles are themselves under test

1. **Reach** — the digest SHALL account for 100% of a part's blocks, asserted as a count equality per corpus fixture. An oracle cannot report on content it never visits, and a coverage number is the only thing that catches an oracle with a hole in it.
2. **Blind-spot pairs** — pairs of documents that *mean* different things (numbering identity changed, borders emptied, a table flattened to loose paragraphs, a style definition gutted) SHALL digest differently.
3. **Teeth** — deliberately strip an element from a saved package and assert the census and digest both fire.
4. **Floor** — every corpus sweep asserts `fixtures.len()` exceeds a pinned floor. An empty glob is the one way a sweeping oracle lies.
5. **Ceilings** — known defects are pinned with exact equality (`== N`, never `<= N`). Fixing a defect without lowering the ceiling fails too; otherwise the headroom hides the next regression.

### D10 — One ledger measures the gap, and it only ratchets

A machine-readable ledger at `crates/betteroffice-docx/tests/scorecard/ledger.json` holds one entry per criterion in the catalogue: status per axis, per-axis target, the gates (test names) proving each claimed status, and known defects. A test recomputes every measurable number — corpus pass rate, census losses, digest reach, required-scenario count, defect counts — and compares with exact equality. The single regeneration gate for the ledger and every golden in the workspace is `GOLDEN_UPDATE=1`; `DL_SNAPSHOT_UPDATE` is retired. A claimed status whose named gate does not exist fails. The gap to 100% is, by construction, the list of (criterion, axis) entries below target plus the defect ledger — enumerable, not estimated.

Alternative rejected: a prose support matrix. Prose drifts; nothing fails when it lies.

### D11 — Layout gates fail hard and grow monotonically

The `goldens.rs` required list becomes the rule, not the exception: a scenario, once required, stays required, and the required count lives in the ledger so shrinking it is a ratchet violation. Corpus fixtures gain pinned layout oracles — page count, drawing count, and the specific geometry that decides a scenario (a page count alone passes while every line breaks in the wrong place). Hand-written fidelity vectors are reviewed against the reference, never regenerated from engine output — regenerating a golden from the code under test turns the oracle into a mirror.

### D12 — Visual gates ride the deterministic rasterizer

`docx-raster` output is byte-deterministic; corpus render goldens therefore compare exact bytes, and a failure writes the actual, expected, and per-pixel diff images as artifacts. Prerequisite: the bytes→measure→layout lowering must be exposed so the pipeline runs `.docx` bytes → `Layout` → `DisplayList` → PNG without a caller-supplied projection — today `render_png` takes a hand-built display list, which blocks visual regression at the root. Any tolerance-based comparison goes through the D6 registry with a declared epsilon.

### D13 — Editing is a fidelity surface with its own oracles

Being able to open a document is not fidelity if the user cannot then edit it the way Word would, so editing gets the same oracle discipline as saving:

1. **Reachability** — every story Word lets a caret enter (body, nested tables, text boxes, headers and footers, footnotes and endnotes, comments), the engine lets a caret enter, and hit testing, caret geometry, and painted glyphs agree on where that caret is.
2. **Edit isolation** — an edit changes only what it means to change. The gate is the *edit-scope fingerprint*: fingerprint the document with the edited subtree replaced by a sentinel, and assert it is unchanged; the digest diff must be exactly the edit's declared footprint, and the census must show no loss.
3. **Undo is exact** — undo after any operation restores a state whose save passes both oracles against the pre-edit save.
4. **Tracked reject is a round trip** — rejecting all tracked changes after an edit session yields the original document's semantic digest. The comparison uses the digest with adjacent identical runs collapsed, not the fingerprint: rejecting an insertion legitimately leaves a split run saying the same text, and Word's own files carry that residue. Choosing the right oracle per property is the rule, not a concession.
5. **Convergence** — concurrent edits converge to one state, and that state passes the same gates. The existing `docx-edit` operation corpora feed this axis.

Alternative rejected: treating editing behavior as UI polish outside the fidelity program. The file Word writes after an edit *is* the behavior; a structurally wrong Enter key produces a structurally wrong document.

### D14 — Word behavior is established by experiment, not by assumption

When ECMA-376 does not decide a behavior, the protocol is: author a probe document, round-trip it through Microsoft Word (desktop or web), check the returned file in as a paired fixture (`<name>-probe.docx` / `<name>-word-roundtrip.docx`), assert the engine's reading is identical on both members, and record in the manifest what Word preserved, normalized, or refused. Engine renders are never labeled as Word references; a Word-comparison artifact carries recorded provenance or it does not exist.

## Placement

| Piece | Location | Why |
| --- | --- | --- |
| Oracles, census, registries | `crates/ooxml-fidelity` | Fingerprint and census are format-neutral (PPTX/XLSX adopt later); the digest is WML-specific in a `wml` module. Dev-dependency of format crates. |
| Corpus | `crates/betteroffice-docx/tests/corpus/` | The facade crate is the one place that composes parse and save. |
| Round-trip and meta gates | `crates/betteroffice-docx/tests/fidelity/` | Same reason; everything runs under plain `cargo test`. |
| Ledger and report | `crates/betteroffice-docx/tests/scorecard/` | Beside the gates that feed it. |
| Fixture generators | `scripts/` | Matches `create-demo-doc.ts`: bun scripts, pinned zip dates. |

Naming: the vocabulary is *structural fingerprint*, *semantic digest*, *element census*, *ledger*. "Canonical" stays reserved for the existing `docx-document-canonical-v1` byte contract in `docx-parse` — the two must not be conflated.

## Risks

| Risk | Mitigation |
| --- | --- |
| The serializer intentionally normalizes and the fingerprint reddens the corpus on day one | D5: declare each normalization deliberately, one review per entry; anything not worth declaring is a bug to fix. |
| Real documents cannot be committed for confidentiality | D7: `betteroffice-redact` intake; the redaction is structure-preserving, so oracles still bite. |
| The digest has holes and passes silently forever | D9 reach + blind-spot pairs + teeth; holes become failing tests, not folklore. |
| The ledger becomes aspirational | D10 claim-gate binding: a status names the test that proves it or the build fails. |
| Corpus growth slows CI | Fixtures are small (tens of KB); the runner is linear over the corpus; budget revisited when measured, not before. |
| Visual goldens churn on every layout fix | Expected and accepted: a golden change is the review surface for a layout change, regenerated only under `GOLDEN_UPDATE=1`. |
