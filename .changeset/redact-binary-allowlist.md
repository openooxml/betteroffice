---
"@betteroffice/rust-crates": patch
---

Scrub unrecognized binary parts during redaction. A part is kept only when it is recognized media or XML; every other part is emptied, and it is removed outright — together with its owned relationship part, that part's exclusive targets and its content-type declaration — when no surviving relationship points at it. The XML rewriter and the scrubber now share one reading of relationship markup, so they cannot disagree about which targets leave the package, and a part is only removed when every surviving relationship resolves to a stored entry.
