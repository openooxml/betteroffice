# Cross-target API parity

`<format>/script.json` runs one call sequence against TS wasm, TS native, Rust and Python.
Each target compares `[{ "op", "result" }]` with `<format>/expected.json`.
An optional step `targets` list selects `"ts"`, `"rust"` and/or `"python"` (default: all); TS always executes every step to produce the golden, while Rust and Python skip excluded steps, retain the golden's version token numbers and fail on unsupported included operations.
The `open` operation takes `{ "fixture": "path/from/repo/root.xlsx" }` to replace the current workbook using the script's calculation settings; `openCollaborative` takes `{ "fixture", "clientId" }` and opens a replica instead.
A step with `"expectError": true` must fail and records `{ "error": message }`, the message each target's call raises.
A step whose TS result differs from the engine's goes in the script's `engineSteps` (docx): Rust and Python run them after `steps` and compare with `<format>/engine-expected.json`, which the Rust test writes under `BETTEROFFICE_PARITY_WRITE=1`.

Normalization spec: Visit steps in script order, arrays in index order and object keys in ascending order. Replace each entire string matching the format's version token pattern (xlsx: 32 lowercase hex digits, pptx: 16 lowercase hex digits, each followed by `-` and a decimal epoch; docx has none) with `<version:N>`, assigning N from 1 on first sight and reusing it for equal tokens across the run. Leave every other value unchanged. In arguments, resolve `<version:N>` to the previously observed token, failing on unknown references. The `save` op records `{ "sha256", "byteLength" }` from the exact saved bytes; docx save instead records `{ "parts": [{ "name", "sha256", "byteLength" }] }` over each part's uncompressed bytes, sorted by name, because package bytes differ in entry timestamps and compression.

To add a step, edit the script and map its op in TS and each selected runner.
Regenerate with `bun scripts/parity/harness.ts xlsx --write` (the wasm engine); add `--preload scripts/native-engine-preload.ts` to run the TS target on the native engine.
`BETTEROFFICE_PARITY_WRITE=1` also enables golden writing in the TS runner.
A TS golden write (`--write` or `BETTEROFFICE_PARITY_WRITE=1`) and the docx Rust engine golden write fail while the environment has `CI` set; compare mode runs as usual there.
All selected targets must pass before accepting a golden.

## docx

Both docx goldens are `{ "scriptSha256", "results": [{ "op", "result" }] }`, where `scriptSha256` is the SHA-256 of the exact bytes of `docx/script.json`.
The TS, Rust and Python runners fail when it does not match the current script, so any script change needs both goldens written again: `bun scripts/parity/harness.ts docx --write` for `expected.json`, then `BETTEROFFICE_PARITY_WRITE=1 cargo test -p betteroffice-docx --test parity` for `engine-expected.json`.
The xlsx and pptx goldens stay plain step arrays.

The read projections compare these fields, with absent values as `null`:

- `paragraphs`, `paragraph` and every nested paragraph: `id`, `style`, `alignment`, `text` and `runs`, each run with `text`, `bold`, `italic`, `underline`, `strike`, `fontSize` (points), `fontFamily` (ascii, else hAnsi), `color` and `styleId`.
- `runHighlights`: per body paragraph, `id` and `highlights`, the `highlight` of each run.
- `tables`: `rows` of `cells` and `text`, each cell with `paragraphs`, nested `tables` and `text`, plus each table's `columnWidths` and `text`.
- `cellMerges`: per table, `rows` of cells, each with `gridSpan` and `vMerge`.
- `sections`: `id`, `pageWidth`, `pageHeight`, `orientation`, the four margins, `paragraphs` and `text`; `headers` and `footers`: `relId`, `kind`, `paragraphs` and `text`.
- `structure`: body paragraph and table counts and the section, header, footer, footnote and endnote counts.
- `readEdited`: the edited paragraph's text and the body story's text, read without a save; the story leaves out table cells. The editing session's own paragraph read is the accepted view (fields as an object character, tracked insertions included, a block content control's paragraph left out) and is not what this step compares.
- `save` and `saveReport`: the parts record above; `saveReport` adds `keptUnchanged` and `structureDropped` of the same save.

The Python binding exposes no run highlight or cell merge properties, so those steps list `"targets": ["ts", "rust"]` and Python skips them.

For a new format, add its script, golden and `scripts/parity/<format>.ts` op table.
Register the table in `harness.ts`, add Rust/Python tests and hook up native CI after its build.
The shared `parity.test.ts` runs every registered table in wasm package tests and native CI.
