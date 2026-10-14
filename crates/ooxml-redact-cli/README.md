# BetterOffice redaction CLI

Redact a local DOCX, XLSX, PPTX, VSDX, or VSTX file:

```sh
cargo run -p betteroffice-redact-cli -- report.docx -o report.redacted.docx
```

The default replaces non-whitespace text with `x`. Use `--random-chars` for independent random letters:

```sh
cargo run -p betteroffice-redact-cli -- report.docx --random-chars -o report.redacted.docx
cargo run -p betteroffice-redact-cli -- japanese.docx --random-chars -o japanese.redacted.docx
```

Random mode preserves Latin, Hiragana, Katakana, and Han scripts automatically, along with whitespace and character counts. Other text symbols become letters from the text node’s dominant supported script. Each replacement uses fresh operating-system randomness; there is no reusable substitution key or seed option. Numeric, date, and formula values use schema-compatible placeholders; typed numeric custom properties use `0`. Style references, relationships, and identifiers keep their structural handling.

Files stay local unless `--share` is supplied. Existing output files and the original input cannot be overwritten.

The Rust API adds `RedactionOptions { random_characters: true }` to `redact_with_options` and `redact_with_report_and_options`. Existing `redact`, `redact_with_report`, and CLI-library `redact_local` calls retain their default behavior; `redact_local_with_options` selects the new mode.
