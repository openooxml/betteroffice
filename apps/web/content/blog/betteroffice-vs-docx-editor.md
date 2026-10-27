---
title: BetterOffice vs docx-editor.dev
description: Milliseconds vs seconds. About 40 ms a page vs about half a second a page. Put the kettle on.
date: 2026-10-24
authors: [BetterOffice maintainers]
tags: [BetterOffice, docx-editor.dev]
---

**577.5 ms vs 55.8 seconds.**

That is the median longest main-thread freeze while each editor opens the same public document. BetterOffice 0.5.0 measures it in milliseconds. docx-editor.dev needs a different unit.

And that is the short part. Five minutes in, docx-editor.dev still had not finished laying the document out. BetterOffice was done in under a minute.

Put the kettle on. This one takes a while.

Versions throughout: BetterOffice 0.5.0, docx-editor.dev 2.32.0, Chromium 149.

## One freeze. 55.8 seconds. And that is just the trailer.

The [public composite document](https://github.com/openooxml/betteroffice/blob/main/scripts/compare/build-public-composite.py) is 15 public documents in one 44.3 MB file. BetterOffice lays it out on 1,263 pages.

Longest single main-thread task, min / median / max, n=6 opens per editor:

- BetterOffice: **483 / 577.5 / 753 ms**
- docx-editor.dev: **47.6 / 55.8 / 63.0 seconds**

Check the units. BetterOffice's worst run: 753 ms. docx-editor.dev's best run: 47.6 seconds.

docx-editor.dev shows the document early. Then, after its open has completed, one task moves in and holds the main thread for 47.6 to 63.0 seconds. Not the whole open. One task. The layout behind it keeps going for minutes.

We gave it more fonts: its fuller documented setup, every default face plus its prefetched pinned font catalog. It gave us **42.0 / 54.3 / 60.1 seconds**. BetterOffice: **453 / 709.5 / 784 ms**. Generous.

## Five minutes in. Still laying out.

We gave both of them five minutes to finish.

BetterOffice laid out all 1,263 pages in 41 to 59 seconds, every run, and never froze the page for longer than 784 ms while doing it. That is about 40 ms a page (32 to 47 ms, derived).

docx-editor.dev, 300 seconds in: page 530 to 597. All six runs. Still laying out. Not finished, not close. That is about half a second a page (0.50 to 0.57 s, derived), and the minute-long freeze is inside those five minutes.

You could make tea while page 600 renders. In 300 seconds, in six runs, page 600 never showed up. BetterOffice was done with page 1,263 by second 59.

Check the units again.

## Review tools: ours ship. Theirs come with section 3.

BetterOffice includes tracked changes, comments and real-time collaboration in its Apache-2.0 packages. Put them in production. No commercial agreement for those features.

docx-editor.dev has the same features in [@docx-editor.dev/pro](https://www.npmjs.com/package/@docx-editor.dev/pro) 2.32.0, under the [EigenPal Pro Evaluation License 1.0](https://unpkg.com/@docx-editor.dev/pro@2.32.0/LICENSE.md). Section 3 is titled **"No Production Use"**. It permits internal, non-production evaluation.

No stopwatch required for this one. Just section 3.

## Word says 23 pages. docx-editor.dev says 20.

The public MIT-licensed [Smart Mailbox final report](https://corpus.betteroffice.dev/smart-mailbox-final-report/source.docx): Word, 23 pages. BetterOffice, 23 pages. docx-editor.dev, 20.

Word puts the four-image panel on page 20. So does BetterOffice. Compared with Word's page, BetterOffice scores **0.984 SSIM**. docx-editor.dev scores **0.697**. (1.0 means identical.)

![Final report page 20 in Word, BetterOffice and docx-editor.dev.](/blog/betteroffice-vs-docx-editor/fid-smart-mailbox-final-report-p20-v232.png)

The [lab notebook](https://corpus.betteroffice.dev/smart-mailbox-lab-notebook/source.docx) from the same project: Word, 37 pages. BetterOffice, 37. docx-editor.dev, 33.

Page 3 holds a meeting record with nested lists. Word draws round bullets. BetterOffice draws round bullets. docx-editor.dev draws hollow squares. Page score: **0.914** for BetterOffice, **0.745** for docx-editor.dev.

![Lab notebook page 3 list in Word, BetterOffice and docx-editor.dev.](/blog/betteroffice-vs-docx-editor/fid-smart-mailbox-lab-notebook-p3-v232-zoom.png)

Milliseconds against seconds. Production against section 3. 23 pages against 20. Round against square.

Check the units. Then count the pages.

## How we measured

Harness and commands: [scripts/compare](https://github.com/openooxml/betteroffice/tree/main/scripts/compare). BetterOffice 0.5.0, docx-editor.dev 2.32.0, Chromium 149, same 6 vCPU VM.

- **Main thread:** the [public composite](https://github.com/openooxml/betteroffice/blob/main/scripts/compare/build-public-composite.py) and its [attribution list](https://github.com/openooxml/betteroffice/blob/main/scripts/compare/public-composite.json) (sources CC-BY-4.0, OGL-UK-3.0, OGL-Canada-2.0, CC0-1.0 and US public domain). Longest main-thread task from the start of open, n=6 opens per editor after one warm-up, alternating order. BetterOffice's window runs until all 1,263 pages are laid out; docx-editor.dev's covers the first 300 s, and in all six runs it was still laying out at 300 s (530 to 597 pages). Its longest task came after its open had completed (38 to 44 s after open began). The font control run used the same windows and n. This measures main-thread responsiveness during open, not total open time.
- **Per-page rates (derived, not separately measured):** docx-editor.dev 300/597 = 0.50 s and 300/530 = 0.57 s per page, averaged over the first 300 s while still laying out. BetterOffice 41/1263 = 32 ms and 59/1263 = 47 ms per page, averaged from the start of open to complete layout.
- **Licences:** [@docx-editor.dev/pro](https://www.npmjs.com/package/@docx-editor.dev/pro) 2.32.0 and the BetterOffice packages. The docx-editor.dev core and React packages are Apache-2.0.
- **Layout:** the [final report](https://corpus.betteroffice.dev/smart-mailbox-final-report/source.docx) and [lab notebook](https://corpus.betteroffice.dev/smart-mailbox-lab-notebook/source.docx) (MIT, [upstream repository](https://github.com/jeongwhanchoi/Smart-Mailbox/tree/e4bf604b55a2f7c465802286f5f289d8364bbeda)). One deterministic capture per editor (n=1), compared with Word 16 page renders by structural similarity (SSIM, 1.0 means identical), scored with the repository's [office-quality compare.py](https://github.com/openooxml/betteroffice/blob/main/scripts/office-quality/compare.py). Each result is one document, not a fidelity average.
