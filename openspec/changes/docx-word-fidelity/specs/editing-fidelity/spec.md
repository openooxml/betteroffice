# editing-fidelity

## ADDED Requirements

### Requirement: Every story Word's caret reaches, ours reaches

Over the corpus, every position a caret can occupy in Word — body text, nested tables, text boxes, headers and footers, footnotes and endnotes, comments — SHALL be reachable, and hit testing, caret geometry, and painted glyph positions SHALL agree: clicking where a character is painted places the caret at that character.

#### Scenario: A nested story accepts the caret

- **WHEN** the user clicks inside a paragraph in a text box inside a table cell
- **THEN** the caret lands in that paragraph at the clicked character, and typing edits that story

#### Scenario: Hit, caret, and paint agree

- **WHEN** the caret is placed at any block's first and last position across the corpus
- **THEN** the hit-test result at the caret's painted rectangle resolves back to the same position

### Requirement: Structural keys produce Word's structures

Enter, Tab, Shift+Tab, Backspace, and Delete SHALL produce the structural outcome Word produces, asserted on the saved package, not only on the in-memory model: Enter splits a paragraph carrying the right properties to the right side; Enter in a list continues the numbering; Tab and Shift+Tab in a list change the level; Backspace at the start of a list item removes the numbering before it removes characters; Tab in a table navigates cells and in the last cell appends a row; Delete and Backspace across a paragraph boundary merge paragraphs the way Word merges them; Backspace cannot destroy table structure from inside a cell.

#### Scenario: Enter in a numbered item

- **WHEN** the caret sits mid-text in a numbered list item and the user presses Enter
- **THEN** the saved file shows two list items in the same numbering, the number sequence advanced by one, and the split text distributed as Word distributes it

#### Scenario: Backspace at cell start

- **WHEN** the caret is at the first position of a table cell and the user presses Backspace
- **THEN** the table structure is unchanged in the saved file

### Requirement: An edit touches only what it means to touch

Every editing operation family SHALL be gated by the edit-scope fingerprint: the document with the edited subtree replaced by a sentinel fingerprints identically before and after the edit. The digest diff SHALL equal the operation's declared footprint, and the census SHALL show no loss.

#### Scenario: A cell edit leaves the host table alone

- **WHEN** text is typed inside a nested table's cell
- **THEN** the outer table's fingerprint with the inner table replaced by a sentinel is unchanged

#### Scenario: A footprint mismatch is a failure

- **WHEN** setting a paragraph property also rewrites a neighboring paragraph's properties
- **THEN** the digest diff exceeds the declared footprint and the gate fails naming the extra path

### Requirement: Undo is exact

After any operation from any family, undo SHALL restore a state whose save passes both oracles against the pre-edit save.

#### Scenario: Undo round-trips

- **WHEN** a paragraph is split and the split is undone
- **THEN** saving yields a package whose fingerprint equals the pre-edit save's and whose digest diff against it is empty

### Requirement: Tracked reject restores the original

Editing under tracked changes and rejecting everything SHALL restore the original document's semantic digest, with adjacent identical runs collapsed before comparison. The fingerprint is deliberately not the oracle here: rejecting an insertion legitimately leaves the run it split standing as two halves saying the same text, and Word's own files carry that residue. Insertions, deletions, and replacements SHALL be exercised at every position class the corpus exhibits — around hyperlinks, note references, fields, and tabs — because offset bookkeeping fails at boundaries, not in the middle of plain text.

#### Scenario: Reject-all after a session

- **WHEN** a tracked session inserts, deletes, and replaces across field and hyperlink boundaries, and every change is rejected
- **THEN** the collapsed digest equals the original document's

### Requirement: Locks and placeholders behave as Word behaves

Content-control locks SHALL refuse exactly the operations Word refuses — `sdtLocked` forbids removing the control, `contentLocked` forbids editing its content, `sdtContentLocked` forbids both, the strictest ancestor winning on each axis. Placeholder text SHALL be a state, not a string: first input replaces the whole prompt and clears the flag; a saved file never carries a placeholder flag over user-typed content.

#### Scenario: Typing beside a prompt

- **WHEN** the user types in a control showing placeholder text
- **THEN** the prompt is replaced wholesale, and the saved control carries the typed text without the placeholder flag

### Requirement: Convergence lands on a faithful state

Concurrent edit histories SHALL converge to a single state, and that state's save SHALL pass the same fidelity gates as a single-author edit. The existing operation corpora in `docx-edit` feed this gate.

#### Scenario: Two authors, one faithful file

- **WHEN** two divergent edit histories from the operation corpus are merged
- **THEN** both replicas reach the same state and its save passes the unedited-gate oracles against itself reopened
