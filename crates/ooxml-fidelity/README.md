# betteroffice-ooxml-fidelity

Fidelity oracles over package bytes. Not published; a dev-dependency of the
format crates' round-trip gates.

Two oracles, neither compensating for the other:

- **Structural fingerprint** — "is this the same tree?" A canonical projection
  of each XML part: prefixes resolve to URIs, attributes sort, insignificant
  whitespace drops under `xml:space` semantics, element order and text stay
  significant. The only forgiven deviations are the reviewed entries in
  `DECLARED_NORMALIZATIONS`.
- **WML semantic digest** — "did a save→reopen keep the meaning?" Per-part
  block records (containment path, attributes, text, nested property tokens,
  generic-subtree fingerprints) plus a structure walk that also covers
  definition parts. `diff_digests` returns paths, never a boolean.

Plus the **element census** (counts by qualified name across every XML part;
`losses` reports what shrank) and the frozen **comparison-mode registry**.

The oracles parse bytes with their own bounded XML reader, never the engine's
parser, so they cannot share its blind spots. Governed by
`openspec/changes/docx-word-fidelity/`.
