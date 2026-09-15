# VSDX synthetic benchmark

Run `bun scripts/vsdx-bench.ts` from the repository root to generate six synthetic
workloads and measure parsing, resolution, evaluation, rendering, and saving.
Each result is a median of five samples after one warm-up.

The renderer has no registered fonts, so its measurements cover fallback text
layout, not font shaping or visual fidelity. These workloads do not replace the
private Visio corpus or Microsoft Visio compatibility checks.

Use `--record` from a clean checkout to replace the baseline with measurements
at the recorded commit. Compare results on the same machine and configuration;
the checked-in historical baseline alone does not establish a performance gain.

## VSDX private-corpus survey

Run `VSDX_EXPLORE_DIR=<dir> cargo run -p betteroffice-vsdx-bench --bin explore`
to sweep every `*.vsdx` and `*.vstx` in `<dir>` (non-recursive) and report parse,
round-trip, resolve, evaluate, render, geometry-row and section-visibility counts.
JSON goes to stdout, a truncated human summary to stderr, and `--json-only`
suppresses the summary. The directory is never committed; the output holds file
names, numeric counts, and bounded category labels only — no cell values,
formulas, shape text, colour literals, author-defined names, or other document
content. Histogram keys are fixed
categories (evaluator error classes, standard Visio function names with
cross-sheet references folded into `<sheet-ref>` and author-defined names in
`<unknown-function>`, standard Geometry row types with author-defined names in
`<unknown-row-type>`, geometry placeholders that name only the failure kind and
row type, section-control names, parse error kinds), capped at 20 entries each,
and parse failures are reported by error kind rather than message. Render
reconciliation is exact at shape level: `shapesPaintedOnly +
shapesPlaceholdered + hidden + unrendered` equals
the shape count, where `unrendered` counts the shapes on pages whose layout
failed and a shape with both a painted and a placeholder primitive counts as
`shapesPlaceholdered`. Primitive counts (`primitivesEmitted`,
`primitivesPainted`, `primitivesPlaceholdered`) are reported separately for
diagnostics and are not shape counts.
