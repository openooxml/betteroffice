# fidelity-corpus

## ADDED Requirements

### Requirement: The corpus is the unit of proof

Fidelity claims SHALL be proven over the corpus at `crates/betteroffice-docx/tests/corpus/`, not over inline XML alone. The corpus SHALL mix provenances deliberately: documents authored by Microsoft Word, LibreOffice, and other real producers alongside deterministic generated fixtures — because the failure modes that matter live in markup real applications write and no synthetic test has a reason to build: stale cached field results, rsid clouds, `mc:AlternateContent` pairs, VML fallbacks.

#### Scenario: Real-producer markup is under test

- **WHEN** the corpus runner executes
- **THEN** at least one fixture authored by each recorded real producer goes through the full round-trip gate

### Requirement: The manifest pins every fixture

`corpus/manifest.json` SHALL record, per fixture: SHA-256 of the checked-in bytes, provenance (`word` | `libreoffice` | `onlyoffice` | `generator` | `redacted`), producing application and version when known, the features exercised, the oracle bound to the fixture, and its evidence status. Aggregate counts SHALL be derived from the entries, never hand-written — a hand-written total silently goes stale. A test SHALL verify every hash and refuse a fixture that has no bound oracle.

#### Scenario: A swapped fixture is caught

- **WHEN** a fixture's bytes change without a manifest update
- **THEN** the SHA-256 check fails

#### Scenario: An oracle-less fixture is refused

- **WHEN** a `.docx` is added to the corpus without a bound oracle
- **THEN** the manifest test fails naming the fixture

### Requirement: Generated fixtures are deterministic and provably current

Every generated fixture SHALL come from a checked-in builder with a pinned zip timestamp and its regeneration command in the builder header. A test SHALL assert the checked-in bytes equal the builder's output, so a builder edit without regeneration — or a hand-edit of the fixture — fails.

#### Scenario: Builder and bytes cannot drift

- **WHEN** a builder is modified and the fixture is not regenerated
- **THEN** the bytes-equal-builder test fails

### Requirement: Fixture naming is uniform

Fixtures SHALL be named by one of: feature description (`table-vmerge-repeated-header.docx`), issue-numbered regression (`issue-NNN-<slug>.docx`), purpose suffix (`-acceptance`, `-showcase`, `-comprehensive`), or evidence pair (`<name>-probe.docx` / `<name>-word-roundtrip.docx`).

#### Scenario: A regression names its issue

- **WHEN** a fixture is added from an engine bug report
- **THEN** its name carries the issue number and a readable slug

### Requirement: The corpus runner gates every fixture

For every manifest entry the runner SHALL execute: the unedited gate (open → save → reopen; fingerprint equality; empty digest diff; zero census loss; byte rules; fixed point) and the edited gate (each edit in the canonical edit set — type a character, split a paragraph, set a paragraph property, set run properties over a range — leaves the census loss-free and produces a digest diff exactly explainable by the edit).

#### Scenario: One fixture, one failure, one name

- **WHEN** any gate fails on any fixture
- **THEN** the failure names the fixture, the gate, and the first difference

### Requirement: Real documents enter through redaction

A confidential document SHALL enter the corpus through `betteroffice-redact`, which preserves structure while scrambling content, and its manifest entry SHALL carry `redacted` provenance with the original producer recorded when known. The `AGENTS.md` repro-file rule gains its destination: an engine PR's repro joins the corpus.

#### Scenario: A production repro becomes a permanent gate

- **WHEN** an engine fix lands with a repro file
- **THEN** the corpus gains the repro (redacted if confidential), its manifest entry, and its oracle, in the same PR
