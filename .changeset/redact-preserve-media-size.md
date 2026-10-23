---
"@betteroffice/rust-crates": patch
---

Redaction can keep image and metafile dimensions, format and approximate byte size with `--preserve-media-size` (`RedactionOptions::preserve_media_size`), using synthetic pixels, so redacted copies load like the originals in benchmarks. Packages whose placeholders would exceed the 512 MiB limit are refused.
