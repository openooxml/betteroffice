# Cross-target API parity

`<format>/script.json` runs one call sequence against TS wasm, TS native, Rust and Python.
Each target compares `[{ "op", "result" }]` with `<format>/expected.json`.
An optional step `targets` list selects `"ts"`, `"rust"` and/or `"python"` (default: all); TS always executes every step to produce the golden, while Rust and Python skip excluded steps, retain the golden's version token numbers and fail on unsupported included operations.
The `open` operation takes `{ "fixture": "path/from/repo/root.xlsx" }` to replace the current workbook using the script's calculation settings.

Normalization spec: Visit steps in script order, arrays in index order and object keys in ascending order. Replace each entire string matching `^[0-9a-f]{32}-[0-9]+$` with `<version:N>`, assigning N from 1 on first sight and reusing it for equal tokens across the run. Leave every other value unchanged. In arguments, resolve `<version:N>` to the previously observed token, failing on unknown references. The `save` op records `{ "sha256", "byteLength" }` from the exact saved bytes.

To add a step, edit the script and map its op in TS and each selected runner.
Regenerate with `BETTEROFFICE_ENGINE=wasm bun scripts/parity/harness.ts xlsx --write`.
`BETTEROFFICE_PARITY_WRITE=1` also enables golden writing in the TS runner.
All selected targets must pass before accepting a golden.

For a new format, add its script, golden and `scripts/parity/<format>.ts` op table.
Register the table in `harness.ts`, add Rust/Python tests and hook up native CI after its build.
The shared `parity.test.ts` runs every registered table in wasm package tests and native CI.
