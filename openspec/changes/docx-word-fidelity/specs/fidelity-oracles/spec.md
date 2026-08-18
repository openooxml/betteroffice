# fidelity-oracles

## ADDED Requirements

### Requirement: Oracles parse bytes, never the typed model

`ooxml-fidelity` SHALL parse XML parts from package bytes with its own namespace-aware, bounded reader, independent of `docx-parse`. Every comparison this capability defines takes original bytes on one side and produced bytes on the other. An oracle built on the code under test inherits its blind spots; this one is blind only to what XML cannot express.

#### Scenario: A model gap does not blind the oracle

- **WHEN** the typed model drops an attribute during parsing and the serializer therefore omits it
- **THEN** the structural fingerprint of the saved part differs from the original and the gate fails, even though the typed structures compare equal

### Requirement: Structural fingerprint

`ooxml-fidelity` SHALL expose a structural fingerprint over a part's XML tree that treats as insignificant: namespace prefix choice (names compare as URI + local name; QName-valued attributes resolve through in-scope bindings), attribute order, quote style, empty-element spelling, and inter-element whitespace that is insignificant under `xml:space` semantics. It SHALL treat as significant: element order, text content including authored whitespace under `xml:space="preserve"` and inside text-bearing WML elements, attribute values, and the set of namespace bindings in scope. Node identities assigned at parse time SHALL NOT enter the fingerprint: two parses of the same bytes assign different ids, and comparing them would report every reopen as a change.

#### Scenario: Lexical noise is invisible

- **WHEN** two parts differ only in prefix choice, attribute order, and empty-element spelling
- **THEN** their fingerprints are equal

#### Scenario: Reordered content is visible

- **WHEN** two parts contain the same elements with two siblings swapped
- **THEN** their fingerprints differ

#### Scenario: Preserved whitespace is significant

- **WHEN** a `w:t` under `xml:space="preserve"` loses a trailing space
- **THEN** the fingerprints differ

### Requirement: Declared normalizations, and only those

The serializer's intentional deviations from input XML SHALL be enumerated in a normalizations registry in `ooxml-fidelity`, one reviewed entry each. The fingerprint applies exactly the registered list before comparing. A difference not covered by a registered entry is a failure, not a shrug.

#### Scenario: An undeclared normalization fails

- **WHEN** the serializer starts reordering children of an element and no registry entry declares it
- **THEN** the unedited round-trip gate fails on the first corpus fixture containing that element

### Requirement: Semantic digest

`ooxml-fidelity::wml` SHALL expose a semantic digest of a package: per story part, an ordered list of block records carrying the containment path from the story root, paragraph text with tabs and hard breaks as characters, paragraph and run property tokens, and whole-subtree fingerprints for generic markup in document order. Property tokens SHALL include their nested children recursively — a bare element name is not a property value: `w:numPr` digests with its `w:numId` and `w:ilvl` values, `w:pBdr` with its edges, `w:tabs` with its stops, and the paragraph mark's run properties with their contents. Within a property container, tokens sort — schema fixes that order, it is not authored meaning; content order is the fingerprint's job. Parts with no story root (`styles.xml`, `numbering.xml`, `settings.xml`) SHALL be digested by the structure walk so a definition part is never covered by no oracle at all.

#### Scenario: A property value change is a digest difference

- **WHEN** a round trip rewrites `w:numPr` from numId 3, ilvl 0 to numId 99, ilvl 5
- **THEN** the digest diff names the paragraph and shows both tokens

#### Scenario: Flattened structure is a digest difference

- **WHEN** a round trip turns a table into the same paragraphs outside any table
- **THEN** the digest diff reports the containment change even though every character survived

#### Scenario: A gutted definition part is a digest difference

- **WHEN** a round trip drops eight of an `abstractNum`'s nine levels
- **THEN** the digest diff names `numbering.xml` and the lost levels

### Requirement: Digest special cases carry their own tokens

The digest SHALL special-case nodes whose meaning a bare block walk under-reports: hyperlinks digest their identity attributes and their inner runs; bookmark starts and ends digest despite carrying no text; drawings digest a structural token including embed relationships and `wp:docPr`; content controls digest their properties separately from their content so an emptied form field cannot digest like a kept one; generic subtrees digest as one fingerprint and are not descended twice; foreign-namespace children of property containers are collected exactly once.

#### Scenario: An emptied control is caught

- **WHEN** a save keeps a content control's shell but drops its content
- **THEN** the digest diff reports the lost content, not equality of shells

### Requirement: Diffs are paths, not booleans

`diff_digests` SHALL return an ordered list of `{path, before, after}` records naming the part, the block, and the field that changed. The failure this oracle exists to catch is a silent drop, and a bare `false` reproduces the silence.

#### Scenario: A loss names its location

- **WHEN** paragraph 12's run properties lose a token in `word/document.xml`
- **THEN** the diff contains one record whose path names that part, that paragraph, and `runProperties`

### Requirement: Two oracles, neither compensates

Every round-trip gate SHALL assert both fingerprint equality and an empty digest diff. A change SHALL NOT weaken one oracle on the grounds that the other passes: the fingerprint cannot see lost meaning in a self-consistent tree, and the digest deliberately tolerates tree residue the fingerprint rejects.

#### Scenario: Fingerprint equal, digest not — still a failure

- **WHEN** a save drops a bookmark in a way that leaves the remaining tree internally consistent across save and reopen
- **THEN** the digest diff is non-empty and the gate fails regardless of the fingerprint

### Requirement: The digest is compared across a real reopen

The after-side of every digest comparison SHALL be computed from a package reopened from the produced bytes, never from the in-memory value that produced them.

#### Scenario: Self-comparison is not a pass

- **WHEN** a serializer bug drops content while the in-memory tree still holds it
- **THEN** the reopened digest differs from the original and the gate fails

### Requirement: Element census

`ooxml-fidelity` SHALL count elements by namespace and local name across every XML part, and report losses — names whose count shrank — between two packages. The census exists to catch drops nobody predicted; it complements the digest, it does not replace it. Where the census is blind — attribute values — identifier multisets (bookmark ids, relationship ids, paragraph ids where present) SHALL have their own guards.

#### Scenario: An unpredicted drop is caught

- **WHEN** an edit-then-save loses one `w:bookmarkStart` anywhere in the package
- **THEN** the census reports `w:bookmarkStart` with a shrunken count

### Requirement: Byte rules

Non-XML parts SHALL be byte-identical through save. XML parts the engine does not model SHALL be byte-identical through save. Modelled XML parts re-emit normalized under both oracles. Save → reopen → save SHALL be byte-identical.

#### Scenario: Media passes through untouched

- **WHEN** a document with images, embedded fonts, and an OLE payload is opened, one paragraph is edited, and the file is saved
- **THEN** every non-XML part's bytes are identical to the input's

#### Scenario: The serializer is a fixed point

- **WHEN** a saved package is reopened and saved again with no edit
- **THEN** the second save's bytes equal the first's

### Requirement: Comparison modes are frozen per artifact

A registry SHALL pin the comparison mode for every fidelity artifact: `exact`, `canonical-exact` with declared ephemera, or `tolerance` with a declared numeric epsilon — the last allowed only for raster comparison. Comparing an unregistered artifact SHALL be an error.

#### Scenario: A tolerance without a declaration refuses

- **WHEN** a test requests a tolerance comparison for an artifact registered as `exact`
- **THEN** the registry refuses instead of comparing loosely
