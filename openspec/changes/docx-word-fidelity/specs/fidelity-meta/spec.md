# fidelity-meta

## ADDED Requirements

### Requirement: The digest reaches everything, and a number proves it

For every corpus fixture, a test SHALL assert that the digest's block count equals the part's actual block count, per story part. An oracle cannot report on content it never visits, and a coverage number is the only thing that catches an oracle with a hole in it.

#### Scenario: A reach hole is a red build

- **WHEN** a new container kind (a nested story, a control level) is parsed but the digest walk does not descend into it
- **THEN** the reach test fails on the first fixture containing it, before any loss occurs there

### Requirement: Blind-spot pairs keep the digest honest

A dedicated test file SHALL hold pairs of documents that mean different things, asserting the digest distinguishes each pair. The set SHALL cover at minimum: numbering identity changed, a border set emptied, tab stops lost, paragraph-mark run properties lost, a table flattened into loose paragraphs, section setup changed, a style definition gutted, a numbering definition gutted, and a content control emptied. Every digest bug found later SHALL add its pair here.

#### Scenario: A meaning change cannot digest equal

- **WHEN** any listed pair is digested
- **THEN** the diff is non-empty and names the changed meaning

### Requirement: The guard has teeth

A test SHALL deliberately strip an element from a saved package and assert that both the census and the digest report the loss. A guard that has never been seen to fire is folklore.

#### Scenario: A synthetic drop trips both nets

- **WHEN** the first paragraph is removed from a saved package's `word/document.xml`
- **THEN** the census reports the shrunken counts and the digest diff names the missing block

### Requirement: A sweeping test proves it swept

Every test that iterates the corpus SHALL first assert the fixture count exceeds a pinned floor. An empty or misrooted glob passes by finding nothing, which is the one way a sweeping oracle lies.

#### Scenario: A moved directory cannot fake a pass

- **WHEN** the corpus directory is renamed and the glob finds zero fixtures
- **THEN** the floor assertion fails before any per-fixture check runs

### Requirement: Known defects are ceilings with exact equality

A known, not-yet-fixed defect class SHALL be pinned as `assert_eq!(defects.len(), N)` — never `<=` — with each counted defect enumerated in the ledger. Fixing one without lowering the ceiling fails too: the headroom would otherwise hide the next regression.

#### Scenario: A fix must claim its fix

- **WHEN** a defect in a pinned class is fixed and the ceiling stays at N
- **THEN** the equality assertion fails, forcing the ceiling — and the ledger — down in the same change

#### Scenario: A regression cannot hide

- **WHEN** a new defect appears while another is fixed, leaving the count at N
- **THEN** the enumerated-defects comparison in the ledger fails even though the count matches
