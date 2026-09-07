# DOCX quality results — 2026-09-07

## What is measured

These historical browser runs used Playwright 1.61.1 with Chromium 149.0.7827.55, a 1400 × 1200 viewport, 150-DPI canvases, and locally reconstructed manifests. The release baseline is the exact `@betteroffice/docx@0.1.0` / `@betteroffice/docx-react@0.1.0` pair from Oxi's lockfile. The candidate includes this PR's engine changes. Both configured runs use the same pinned CDN font assets. These totals precede the final header-path and error-reporting review fixes, which have separate regression tests. Candidate font-load records are retained locally with each result; the initial release runs predate that instrumentation. The subsequent harness privacy check disables the older automatic Google Fonts lookup as well.

The current source baseline is `0f474aa53e2d6ac05dd11cce7eff78901f480273`. Development documents were selected separately from the evaluation documents.

## Confirmed regressions fixed

| Check | Before | After |
| --- | --- | --- |
| First Japanese evaluation document, no font provider, current source baseline | 30-second timeout | 2 pages in 4.94 seconds |
| Original twelve-paragraph header fixture, CDN fonts | 12 pages; header decoration missing | 1 page; decoration visible at its authored coordinates |
| Original alternate-main-part fixture, CDN fonts | Blank page | All twelve paragraphs and the header restored; full and selective saves preserve the actual part |
| Japanese educational development document, local fonts | 194 pages | 12 pages |
| Japanese evaluation corpus, CDN fonts | Published release: 50/50 completed, 305 pages | Candidate: 50/50 completed, 200 pages |
| English evaluation corpus, CDN fonts | Published release: 50/50 completed, 1,042 pages | Candidate: 50/50 completed, 663 pages |
| Candidate CDN font loads, both languages | — | 1,283 successful loader calls, zero failures |

All 200 candidate Japanese pages were exported. Sixteen Japanese evaluation documents changed page count. Fewer pages alone are not an accuracy score: the header regression has a geometric test and before/after screenshots, while the corpus still needs Word reference images. The educational document's cached `docProps/app.xml` page count is not used as ground truth.

The published release also timed out on the first Japanese document without fonts at a 45-second parent deadline. Supplying CDN fonts completed that document in 7.07 seconds and completed all 50 Japanese documents. The candidate's cache fix additionally removes the unconfigured-font preflight stall.

The full unconfigured Japanese control also completed 50/50 documents and exported 142 pages. These renders use approximate fallback metrics and are a completion check. They are not a substitute for font-configured quality measurements.

All 863 candidate pages with CDN fonts were exported. English release exports include two long documents (246 and 481 pages) that exceeded an initial 90-second deadline and completed under a 600-second deadline. Their successful exports took about 450 and 255 seconds. Candidate English runs also used 600-second deadlines. These timings include exporting every page and do not demonstrate an engine speedup; concurrency differed between runs.

One valid English package uses `word/document2.xml` through its root relationship. The release rendered a blank page; the final candidate restores its 1,061 text characters across fourteen blocks. The import fix and conservative shape-placement guard were applied after the full browser runs: native layout hashes confirmed that 91 documents were unchanged, and all nine affected documents were rendered again in the browser. The affected Japanese files were also rerun without fonts. The result records merge those completed runs and explicitly include all 50 selected files per language.

Per-document results, manifests, layout hashes, PDFs, screenshots, and downloaded corpus files remain local under ignored `.source/docx-quality`. The PR adds no dataset or document binaries; Rust regression fixtures are generated in memory. Body shapes with square, tight, through, or top-and-bottom wrapping retain their existing flow approximation; this PR does not implement shape exclusion zones.

## Published Oxi numbers and target

At [Oxi commit e3b653d](https://github.com/Ryujiyasu/oxi/tree/e3b653d81394d49d7a5c9b649e595b488eebd299), the English leaderboard reports:

| Engine | Common-page SSIM | Page-penalized SSIM | Exact page-count matches |
| --- | ---: | ---: | ---: |
| ONLYOFFICE | 0.908 | 0.903 | 42/48 |
| Oxi | 0.903 | 0.903 | 48/48 |
| LibreOffice | 0.892 | 0.881 | 43/48 |
| SILURUS | 0.791 | 0.741 | 35/48 |
| BetterOffice | 0.761 | 0.619 | 26/48 |

Reaching the listed fourth-place scores would require +0.030 common-page SSIM and +0.122 page-penalized SSIM. Those are targets, not measured gains from this PR.

**Comparable final SSIM and leaderboard position remain unavailable.** Searches of Oxi's GitLab repository, GitHub mirror/fork, artifacts, and benchmark-directory history did not locate its original frozen manifests or Word PNGs. Its Windows Word oracle is Microsoft 365 16.0.20131.20154. Local macOS Word 16.112.3 exports now work when inputs and outputs are staged inside Word's own container; two synthetic fixture exports verified the approach. The [small Office harness](../office-quality/README.md) records the permission learning and creates separately labeled Mac references. All 50 English selections are valid packages. The original two Word-failure IDs are unknown, so the original 48-document scoring denominator cannot be reconstructed from the public artifacts alone.

To reproduce Oxi's score, supply the original frozen manifests, Word page images, and the two English Word-failure IDs/reasons. Alternatively, rerun all compared engines against a newly generated, consistently labeled reference set. The local comparator rejects incomplete render records and penalizes missing/extra pages. No top-three or top-four claim is made here.

## Validation

The separate public demo check uses the existing `betteroffice-demo.docx` (SHA-256 `5b272a248fe899d3b90259297dc20158b1bdfc7c0e1d653c7e29985869e17c51`). Word 16.112.3 exported two pages at 18:48:23 UTC on 2026-09-07. Both published 0.1.0 and the final candidate, with identical CDN fonts and Google Fonts disabled, produced two pages and **0.7710525848 common/page-penalized SSIM**. Page scores were 0.6835701611 and 0.8585350085. All images were 1275 × 1650 at 150 DPI; no resize was used. This demo shows no score gain. Its source, Word PNGs, and provenance/score JSON are [published separately in R2](../../README.md#visual-fidelity); the downloaded evaluation corpus remains local.

- 699 Rust tests passed across the DOCX editing, layout, and parsing crates; one existing test remains ignored.
- Strict Clippy checks passed for all three crates and all their targets.
- 44 font, measurement, and React hook tests passed after the compatibility check.
- DOCX, DOCX React, and demo TypeScript checks passed.
- Font, DOCX, and DOCX React package builds passed.
- Four Python comparison checks passed for page penalties, explicit resizing, source matching, incomplete-render rejection, and missing page filenames.
- The local Word reference/candidate pipeline exported one synthetic page on each side and scored 0.9728 SSIM with explicit resize-to-match scoring. This verifies the harness, not a corpus rank.
- The final Japanese privacy check exported two pages in 5.39 seconds with no errors: five pinned font GETs, no request bodies, cookies, referrers, or other external requests.

The font CDN build check rejects binary asset imports and imports of the optional CJK package. Font tests verify real SFNT/CFF payloads, pinned Latin/Japanese URLs, shared buffers, HTTP failure recovery, truncated-payload rejection, and bounded fetch signals.
