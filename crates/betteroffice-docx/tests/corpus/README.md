# DOCX corpus tests

`manifest.json` lists every fixture with its SHA-256, provenance, features and
oracle name. Tests require at least two fixtures, matching hashes, and no
unlisted files except Word owner files beginning with `~$`.

Each fixture is opened and saved, then compared with the
[package oracles](../../../ooxml-fidelity/README.md). Every finding must match
`expected/<fixture>.findings.txt` exactly. Reopening and saving the result must
produce identical package bytes; unstable parts are findings. These checks
compare package contents and do not establish rendering agreement with Word.

A source root that binds a standard WML prefix such as `w16` to another namespace
is saved with the serializer's fixed Word binding, which can change
`mc:Choice/@Requires` meaning; the oracles report this loss.

Run `cargo test -p betteroffice-docx --test corpus`. To regenerate findings,
run `GOLDEN_UPDATE=1 cargo test -p betteroffice-docx --test corpus`, then review
the complete diff. Updating findings is explicit and permits changed results.

`bun test scripts/corpus-fixtures.test.ts` regenerates both fixtures in temporary
directories and requires byte equality with the checked-in files. It also
checks the manifest hashes and that the demo app ships the same demo fixture.
