# visual-fidelity

## ADDED Requirements

### Requirement: The pipeline runs from bytes to pixels

The bytes→measure→layout lowering SHALL be exposed so the engine renders `.docx` bytes to `Layout`, `DisplayList`, and PNG with no caller-supplied projection. Until it is, visual regression cannot exist: a rasterizer that only accepts hand-built display lists tests the rasterizer, not the document.

#### Scenario: A document renders end to end

- **WHEN** corpus fixture bytes are passed to the render entry point with registered fonts
- **THEN** a PNG per page is produced with no hand-supplied measurement or layout input

### Requirement: Corpus render goldens compare exact bytes

Selected corpus fixtures SHALL have per-page PNG goldens. Rendering is byte-deterministic, so comparison is byte equality — no silent tolerance. Any tolerance-based comparison that proves necessary goes through the comparison-mode registry with a declared epsilon.

#### Scenario: A one-pixel change is a failure

- **WHEN** a paint change moves any pixel on a golden page
- **THEN** the golden test fails; acceptance is an explicit regeneration under `GOLDEN_UPDATE=1`, reviewed as the visual diff it is

### Requirement: A visual failure shows itself

A failing render golden SHALL write the actual image, the expected image, and a per-pixel difference image as test artifacts. A byte-inequality message without pictures forces the reviewer to reproduce locally before they can even see the regression.

#### Scenario: The diff is one open away

- **WHEN** a render golden fails in CI
- **THEN** the three images are available as artifacts, named by fixture and page

### Requirement: Word-reference evidence carries provenance or does not exist

An engine render SHALL never be labeled as a Word reference. A comparison against Word SHALL exist as a record in the fixture's manifest entry naming the Word variant and version, the date, the source file, and what was observed. Images belong in the pull request that makes the claim, never in the repository: an image no test compares is prose that costs bytes. The ledger's paint axis reaches `golden` on engine determinism alone; it reaches Word-verified evidence only through this protocol.

#### Scenario: An unlabeled screenshot is rejected

- **WHEN** a change adds a "matches Word" claim backed by an image without recorded provenance
- **THEN** review rejects the claim; the image is either labeled engine output or given its manifest record
