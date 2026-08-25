---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Make editing cost page-local on large documents: fix the section count that disabled the resident fast path, coalesce per-cluster text primitives into per-run primitives, merge selection rects per line, bound display rebuilds to damaged pages, reuse retained measures on structural relayouts, and stop shipping the measured arena to the host.
