---
title: BetterOffice vs Univer
description: Week 54 in January, the number 1 in fourth place, 9:00 AM at #VALUE!. Welcome to Univer Standard Time.
date: 2026-10-24
authors: [BetterOffice maintainers]
tags: [BetterOffice, Univer]
---

**Monday, January 4, 2021. Week 1 in BetterOffice. Week 54 in Univer.**

It had a big weekend.

Welcome to Univer Standard Time, where the number 1 finishes fourth, 9:00 AM is #VALUE! and BetterOffice's slowest cell edit beats Univer's fastest. No adjectives survived contact with these cells.

Versions throughout: BetterOffice 0.5.0, Univer presets 1.2.0, Chromium 149.

## Bring the .xlsx. BetterOffice opens it.

BetterOffice opens .xlsx bytes directly in the browser. That is the whole workflow.

Univer's open-source packages load IWorkbookData snapshots, not .xlsx bytes. An application supplies its own converter or uses Univer's [documented import](https://docs.univer.ai/guides/sheets/features/import-export), which runs through its Pro exchange packages and a conversion backend. For these tests we handed Univer pre-converted snapshots from our own adapter. BetterOffice got the file.

## Week 54, four days into January

January 4, 2021 is a Monday. ISO 8601 and Microsoft's ISOWEEKNUM documentation both put it in week 1.

`=ISOWEEKNUM(DATE(2021,1,4))`

- BetterOffice: **1**
- Univer: **54**

That was not a one-off. Across n=10 dates in the [ISO week workbook](https://github.com/openooxml/betteroffice/blob/main/scripts/compare/sheets/edge/formula_probe.py), checked against Microsoft's definition and an independent ISO calendar, BetterOffice matches **all 10**. Univer differs on **8**.

![BetterOffice: =ISOWEEKNUM(DATE(2021,1,4)) returns 1.](/blog/betteroffice-vs-univer/sheets-formula-isoweeknum-bo-v120.png)

![Univer: =ISOWEEKNUM(DATE(2021,1,4)) returns 54.](/blog/betteroffice-vs-univer/sheets-formula-isoweeknum-univer-v120.png)

## The number 1 finishes fourth

Microsoft's documentation says RANK.EQ ignores non-numeric values in the reference. In the [ranking workbook](https://github.com/openooxml/betteroffice/blob/main/scripts/compare/sheets/edge/formula_probe.py), E1:E3 hold 1, 2, 3. E4:E6 hold formulas returning `""`. Nothing.

`=RANK.EQ(1,E1:E6,1)` returns **1** in BetterOffice and **4** in Univer. RANK does the same.

Three empty results entered the race. The smallest number in the range came fourth.

![BetterOffice RANK.EQ result.](/blog/betteroffice-vs-univer/sheets-formula-rankref-eq-bo-v120.png)

![Univer RANK.EQ result.](/blog/betteroffice-vs-univer/sheets-formula-rankref-eq-univer-v120.png)

## Your quarter hour has been declined

`=MROUND(0.378,"0:15")` rounds a time to the nearest quarter hour.

BetterOffice returns **0.375 (9:00 AM)**. Univer returns **#VALUE!**.

Nothing says 9:00 AM like #VALUE!. Five of five such expressions in the [time-rounding workbook](https://github.com/openooxml/betteroffice/blob/main/scripts/compare/sheets/edge/formula_probe.py): same result.

![BetterOffice MROUND result.](/blog/betteroffice-vs-univer/sheets-formula-mround-time-bo-v120.png)

![Univer MROUND result.](/blog/betteroffice-vs-univer/sheets-formula-mround-time-univer-v120.png)

## Awkward podium

Scroll a synthetic 30,000-row workbook on a 6 vCPU [Kyora](https://kyora.sh) VM. Frames per second, min / median / max:

- BetterOffice: **29.6 / 37.7 / 41.4**
- Univer: **17.0 / 19.0 / 19.8**

BetterOffice's worst session: 29.6. Univer's best: 19.8.

Now edit some cells in those same sessions. From pressing Enter to the first frame showing the value, 24 edits per editor, min / median / max:

- BetterOffice: **33 / 54 / 93 ms**
- Univer: **149 / 236.5 / 3,539 ms**

BetterOffice's slowest edit: 93 ms. Univer's fastest: 149 ms. Univer's slowest: 3.539 seconds.

Read it again. We did.

Week 54. Fourth place for the number 1. A quarter hour at #VALUE!. Up to 3,539 ms to show an edited cell. Welcome to Univer Standard Time.

## How we measured

Harness and commands: [scripts/compare](https://github.com/openooxml/betteroffice/tree/main/scripts/compare). BetterOffice 0.5.0, Univer presets 1.2.0, Chromium 149.

- **Import:** BetterOffice opens .xlsx bytes. Univer opened IWorkbookData snapshots pre-converted by our ExcelJS-based adapter, so open time is not compared.
- **Formulas:** the [ISO week](https://github.com/openooxml/betteroffice/blob/main/scripts/compare/sheets/edge/formula_probe.py) (n=10 dates), [ranking](https://github.com/openooxml/betteroffice/blob/main/scripts/compare/sheets/edge/formula_probe.py) (n=2 failing fixtures, one per function, plus one passing control holding only the numbers, which returns 1 in both; cell references, not array constants) and [time-rounding](https://github.com/openooxml/betteroffice/blob/main/scripts/compare/sheets/edge/formula_probe.py) (five of five expressions) workbooks. Typed inputs and no cached values, so the conversion cannot change the results. Each expected value checked against Microsoft's documentation and an independent calculation; each Univer value traced to the function's implementation in @univerjs/engine-formula 1.2.0.
- **Scrolling and edits:** a 30,000-row workbook from the [scale generator](https://github.com/openooxml/betteroffice/blob/main/scripts/compare/sheets/gen.mjs), same 6 vCPU [Kyora](https://kyora.sh) VM, n=6 interleaved measured sessions per editor after one warm-up. 5 s of mouse-wheel input from the middle row (one 100 px wheel event every 50 ms); frames are the ones Chromium's DevTools screencast (`Page.screencastFrame`) delivered during those 5 s. 4 cell edits per session (24 per editor), timed from Enter to the first frame showing the value.
