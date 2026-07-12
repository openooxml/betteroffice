/**
 * Golden regression corpus.
 *
 * Each scenario pairs a synthetic display list with a viewport state and the
 * visible-cell range the pure math derives from it. The committed
 * `golden/<name>.json` files freeze the CURRENT output verbatim — they do not
 * judge correctness, only pin it so the display-list + viewport seam can prove
 * byte-identical output across refactors and the eventual Rust port. See
 * README.md.
 */

import type { DisplayList } from '../types';
import type { VisibleCells, ViewportState } from '../../viewport/index';
import { visibleCells } from '../../viewport/index';
import {
  fillRect,
  line,
  makeGridDisplayList,
  makeUniformGridOffsets,
  makeViewport,
  text,
} from './factories';

/**
 * The serializable snapshot pinned per scenario: the frame the renderer would
 * paint plus the virtualization decision that selects it.
 */
export interface GoldenSnapshot {
  displayList: DisplayList;
  viewport: ViewportState;
  visible: VisibleCells;
}

/**
 * One pinned scenario.
 */
export interface GoldenScenario {
  /** kebab-case name; also the golden file basename. */
  name: string;
  /** one-line description of what this pins. */
  pins: string;
  build(): GoldenSnapshot;
}

const COL_W = 80;
const ROW_H = 20;

function plainGridNoFreeze(): GoldenScenario {
  return {
    name: 'plain-grid-no-freeze',
    pins: 'a small unscrolled grid with no frozen panes shows its leading tracks',
    build() {
      const viewport = makeViewport({ width: 320, height: 120 });
      const { colOffsets, rowOffsets } = makeUniformGridOffsets(20, 40, COL_W, ROW_H);
      return {
        displayList: makeGridDisplayList(4, 6, COL_W, ROW_H),
        viewport,
        visible: visibleCells(viewport, colOffsets, rowOffsets),
      };
    },
  };
}

function scrolledPastFirstScreen(): GoldenScenario {
  return {
    name: 'scrolled-past-first-screen',
    pins: 'scrolling clips leading tracks and reveals a later window',
    build() {
      const viewport = makeViewport({ width: 320, height: 120, scrollX: 250, scrollY: 130 });
      const { colOffsets, rowOffsets } = makeUniformGridOffsets(20, 40, COL_W, ROW_H);
      return {
        displayList: makeGridDisplayList(4, 6, COL_W, ROW_H),
        viewport,
        visible: visibleCells(viewport, colOffsets, rowOffsets),
      };
    },
  };
}

function frozenHeaderRowAndCol(): GoldenScenario {
  return {
    name: 'frozen-header-row-and-col',
    pins: 'one frozen row + column pin while the scrolled body advances past them',
    build() {
      const viewport = makeViewport({
        width: 320,
        height: 120,
        scrollX: 200,
        scrollY: 100,
        frozenRows: 1,
        frozenCols: 1,
      });
      const { colOffsets, rowOffsets } = makeUniformGridOffsets(20, 40, COL_W, ROW_H);
      return {
        displayList: makeGridDisplayList(4, 6, COL_W, ROW_H),
        viewport,
        visible: visibleCells(viewport, colOffsets, rowOffsets),
      };
    },
  };
}

function clippedAlignedCellText(): GoldenScenario {
  return {
    name: 'clipped-aligned-cell-text',
    pins: 'text commands carry clip rects and per-cell horizontal alignment',
    build() {
      const viewport = makeViewport({ width: 240, height: 60 });
      const { colOffsets, rowOffsets } = makeUniformGridOffsets(3, 2, COL_W, ROW_H);
      const displayList: DisplayList = {
        width: 3 * COL_W,
        height: 2 * ROW_H,
        commands: [
          text(2, 16, 'left', { clip: { x: 0, y: 0, w: COL_W, h: ROW_H }, align: 'left' }),
          text(COL_W + 40, 16, 'mid', {
            clip: { x: COL_W, y: 0, w: COL_W, h: ROW_H },
            align: 'center',
          }),
          text(3 * COL_W - 2, 16, 'right long overflow', {
            clip: { x: 2 * COL_W, y: 0, w: COL_W, h: ROW_H },
            align: 'right',
          }),
        ],
      };
      return { displayList, viewport, visible: visibleCells(viewport, colOffsets, rowOffsets) };
    },
  };
}

function styledCellsFillFontBorder(): GoldenScenario {
  return {
    name: 'styled-cells-fill-font-border',
    pins: 'fill + font facets (bold/italic/underline/strike) + solid/dashed/double borders serialize additively',
    build() {
      const viewport = makeViewport({ width: 2 * COL_W, height: 2 * ROW_H });
      const { colOffsets, rowOffsets } = makeUniformGridOffsets(2, 2, COL_W, ROW_H);
      const clip = (c: number, r: number): { x: number; y: number; w: number; h: number } => ({
        x: c * COL_W,
        y: r * ROW_H,
        w: COL_W,
        h: ROW_H,
      });
      const displayList: DisplayList = {
        width: 2 * COL_W,
        height: 2 * ROW_H,
        commands: [
          // a filled header cell with bold white text over it.
          fillRect(0, 0, COL_W, ROW_H, '#4472c4'),
          text(2, 14, 'Header', {
            clip: clip(0, 0),
            align: 'left',
            color: '#ffffff',
            bold: true,
          }),
          // an italic + underlined + struck currency cell, right-aligned.
          text(2 * COL_W - 2, 14, '1,250.00', {
            clip: clip(1, 0),
            align: 'right',
            italic: true,
            underline: true,
            strike: true,
            fontFamily: 'Calibri',
          }),
          // border edges: solid bottom, dashed right, double bottom on the row.
          line(0, ROW_H, COL_W, ROW_H, 2, '#1f3864'),
          line(COL_W, 0, COL_W, ROW_H, 1, '#1f3864', 'dashed'),
          line(0, 2 * ROW_H, 2 * COL_W, 2 * ROW_H, 1, '#1f3864', 'double'),
        ],
      };
      return { displayList, viewport, visible: visibleCells(viewport, colOffsets, rowOffsets) };
    },
  };
}

/**
 * The full corpus. Order is stable so the suite runs deterministically.
 */
export const corpus: GoldenScenario[] = [
  plainGridNoFreeze(),
  scrolledPastFirstScreen(),
  frozenHeaderRowAndCol(),
  clippedAlignedCellText(),
  styledCellsFillFontBorder(),
];
