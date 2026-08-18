# DOCX Word fidelity

## Why

The engine's fidelity claims currently rest on tests that cannot see the failures they exist to prevent.

**Nothing measures loss on real documents.** The `crates/` tree contains four binary OOXML fixtures in total — `docx-edit/tests/fixtures/footnote-anchor.docx`, `ooxml-text/tests/fixtures/line-spacing-baseline.docx`, `pptx-parse/tests/fixtures/chart-deck.pptx`, `betteroffice-xlsx/tests/fixtures/hidden-dimensions.xlsx`. Every other DOCX-shaped test builds its package inline from string-literal XML. A parser or serializer bug that only manifests on markup Word actually produces — stale cached field results, `mc:AlternateContent` pairs, rsid clouds, VML fallbacks — has no fixture to manifest on.

**Round-trip fidelity is asserted on the typed structure, not on the package.** `crates/betteroffice-docx/tests/document.rs` `opens_edits_saves_and_reopens_typed_structure` compares `reopened.structure()` to the original structure. A save that drops an element the structure projection does not carry passes. The single byte-level save assertion, `an_unreadable_chart_part_survives_a_save_byte_for_byte`, covers one unreadable part. Save-fidelity tests exist only as per-bug additions (`saving_keeps_shape_text_box_bodies_and_character_unit_indents` and siblings), each guarding the one shape it was written for.

**`docx-parse` — the largest DOCX crate — has no fixture-driven surface.** It has no `tests/` directory; no test in the workspace opens a real-world document, saves it untouched, and diffs the packages part by part.

**The layout gate fails soft.** `crates/docx-layout/tests/goldens.rs` runs 14 scenarios but requires only 2; the other 12 may silently degrade to `UNSUPPORTED` without reddening CI. Golden regeneration is inconsistent: `DL_SNAPSHOT_UPDATE=1` for display lists and header/footer bands, `GOLDEN_UPDATE=1` for raster PNGs, no gate at all for the hand-maintained `.golden.json` files and the inline chart snapshot.

**Editing parity is untracked.** `docx-edit` has operation and convergence corpora, but nothing states which constructs a user can actually edit, whether the caret can reach every story Word's caret reaches, whether Enter, Tab, and Backspace produce the structures Word produces, or whether an edit leaves the rest of the document untouched. An editor whose Enter key builds the wrong structure writes the wrong file; that is a fidelity failure, not a UI detail.

**The gap is unmeasured.** There is no artifact stating which WordprocessingML features the engine preserves, models, edits, lays out, or paints, at what confidence, with which known defects. "How far from 100% Word fidelity?" has no answer today, so it cannot ratchet.

**The intake pipe exists but is unwired.** `AGENTS.md` requires a repro file on every engine PR, and `betteroffice-redact` produces structure-preserving, shareable versions of confidential documents — yet no committed corpus receives them.

## What Changes

- **`crates/ooxml-fidelity`** — a new dev-facing crate holding the instruments: a neutral XML tree reader, the structural fingerprint, the WML semantic digest with path-diffing, the element census, and the comparison-mode registry. The oracles parse saved bytes independently of `docx-parse`, so they cannot share the parser's blind spots.
- **Round-trip gates in `betteroffice-docx`** — every fixture: open → save → reopen, assert fingerprint equality, empty digest diff, zero census loss, byte identity for non-XML and unmodelled parts, and save→reopen→save byte-level fixed point.
- **A golden corpus** at `crates/betteroffice-docx/tests/corpus/` — Word/LibreOffice-authored documents plus deterministic generated fixtures, pinned by a SHA-256 manifest in which every fixture is bound to a named oracle. Real-document intake flows through `betteroffice-redact`.
- **Acceptance and comprehensive fixtures** — a generated document covering every supported run and paragraph property that self-validates against the parser's supported-property inventory, and a Word-authored comprehensive document with section-numbered oracles.
- **Editing-fidelity gates** — every story reachable by the caret with hit testing, caret geometry, and paint in agreement; Word-parity semantics for the structural keys (Enter, Tab, Backspace, Delete across boundaries) in lists, tables, and nested stories; edit isolation proven by an edit-scope fingerprint; exact undo; tracked reject restoring the original digest; convergence feeding the same gates.
- **Meta-oracles** — tests that test the oracles: digest reach must cover 100% of blocks, blind-spot pairs must digest differently, a deliberately broken document must trip the guard, the corpus glob must have found the corpus, and known defects are pinned by exact-equality ceilings.
- **A fidelity scorecard** — a criteria catalogue spanning WordprocessingML, a machine-readable ledger scoring each criterion on six axes against a target, a claim-gate binding test (a claimed status must name the test that proves it), a generated human-readable report, and a CI ratchet: measured numbers may only improve.
- **Layout and visual hardening** — one regeneration gate (`GOLDEN_UPDATE=1`) for every golden, a required-scenario list that only grows, per-corpus-fixture page and geometry pins, exposure of the bytes→measure→layout lowering so the pipeline runs `.docx` bytes to PNG end to end, and corpus render goldens on the byte-deterministic rasterizer.
- **A Word-evidence protocol** — when ECMA-376 is silent, behavior is established by experiment: a probe document round-tripped through Microsoft Word, the returned file checked in as a paired fixture, the observed preservations and normalizations recorded.

## Impact

- New crate `crates/ooxml-fidelity` (fingerprint and census format-neutral; digest WML-specific in a `wml` module).
- `crates/betteroffice-docx`: new `tests/corpus/` (fixtures + manifest), `tests/fidelity/` (round-trip and meta gates), `tests/scorecard/` (ledger + generated report).
- `crates/docx-edit`: editing-fidelity gates (edit-scope fingerprint, undo, tracked reject) as dev-dependencies on `ooxml-fidelity`.
- `crates/docx-layout`: `goldens.rs` required-list hardening; `DL_SNAPSHOT_UPDATE` replaced by `GOLDEN_UPDATE`; the hand-edited `.golden.json` and inline chart snapshot gain the same gate.
- `crates/docx-raster`: corpus render goldens; failure emits a diff artifact.
- `scripts/`: deterministic fixture generators (acceptance document, probe documents).
- `AGENTS.md`: the repro-file rule gains its destination — the repro joins the corpus, redacted when confidential.
- CI: no new jobs; every gate runs inside the existing `cargo test`.
