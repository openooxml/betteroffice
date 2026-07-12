# Golden display-list corpus

The safety net for the display-list + viewport seam. It pins the current output
of the viewport math and the synthetic display-list factories so refactors — and
the eventual Rust/WASM port — can prove **byte-identical** output.

The goldens capture **current** behavior verbatim. They are not a statement that
the output is _correct_ — only that it must not change unintentionally.

## What's here

- `factories.ts` — DOM-free builders for display lists (fills, lines, clipped
  aligned text, gridlines) and viewport states / track offsets.
- `corpus.ts` — named `{ name, pins, build() }` scenarios. `build()` returns a
  `{ displayList, viewport, visible }` snapshot; `visible` is computed by the
  real `visibleCells` math, so the goldens exercise virtualization for real.
- `serialize.ts` — deterministic canonicalizer: sorted keys, `undefined`
  dropped, every number rounded to `GOLDEN_PRECISION` (3) decimals, `-0`
  collapsed.
- `golden/<name>.json` — one committed baseline per scenario.
- `golden.test.ts` — builds each scenario, serializes, asserts equality.

## Scenarios

| Scenario                    | Pins                                                              |
| --------------------------- | ---------------------------------------------------------------- |
| `plain-grid-no-freeze`      | a small unscrolled grid with no frozen panes                     |
| `scrolled-past-first-screen`| scrolling clips leading tracks and reveals a later window        |
| `frozen-header-row-and-col` | a frozen row + column pin while the scrolled body advances       |
| `clipped-aligned-cell-text` | text commands with clip rects and per-cell horizontal alignment  |

## Regenerating

A normal test run never writes files. To (re)write the goldens, gate on the env
var (from `packages/core`):

```bash
GOLDEN_UPDATE=1 bun test src/display-list/__golden__
```

Only regenerate when a diff is **intended and reviewed**. A nonzero diff during
a refactor means **stop and investigate**: either an intended equivalence
(regenerate deliberately and review the JSON) or a regression (fix the code).
