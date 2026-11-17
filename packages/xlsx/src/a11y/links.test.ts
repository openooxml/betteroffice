import { describe, expect, it } from 'bun:test';
import { buildA11yGrid } from './index';
import type { A11yStrings } from './index';
import type { DisplayList, HyperlinkRegion } from '../display-list/types';

const strings: A11yStrings = {
  gridLabel: 'Spreadsheet grid',
  cellLabel: '{address}, {value}',
  cellLabelSelected: '{address}, {value}, selected',
  emptyCellLabel: '{address}, empty',
  emptyCellLabelSelected: '{address}, empty, selected',
  rowHeaderLabel: 'Row {row}',
  columnHeaderLabel: 'Column {column}',
};

function frame(hyperlinks: HyperlinkRegion[], grid?: Partial<NonNullable<DisplayList['grid']>>): DisplayList {
  return {
    width: 160,
    height: 40,
    commands: [],
    grid: { startRow: 0, startCol: 0, colOffsets: [0, 80, 160], rowOffsets: [0, 20, 40], ...grid },
    hyperlinks,
  };
}

const at = (link: Omit<HyperlinkRegion, 'top' | 'left' | 'bottom' | 'right'>, row = 0, col = 0): HyperlinkRegion =>
  ({ top: row, left: col, bottom: row, right: col, ...link });

const hrefs = (list: DisplayList) =>
  buildA11yGrid(list, null, 'Sheet1', strings).rows.map((row) => row.cells.map((cell) => cell.href ?? null));

describe('buildA11yGrid links', () => {
  it('exposes safe external targets and places in the workbook', () => {
    expect(hrefs(frame([
      at({ externalTarget: 'https://example.com/a' }),
      at({ externalTarget: 'mailto:a@example.com' }, 0, 1),
      at({ location: "'Other sheet'!B2" }, 1, 0),
      at({ location: '#Sheet1!A1' }, 1, 1),
    ]))).toEqual([
      ['https://example.com/a', 'mailto:a@example.com'],
      ["#'Other sheet'!B2", '#Sheet1!A1'],
    ]);
  });

  it('never exposes a target the click path refuses', () => {
    for (const externalTarget of [
      'javascript:alert(1)',
      ' javascript:alert(1)',
      'java\tscript:alert(1)',
      'data:text/html,<script>alert(1)</script>',
      'vbscript:msgbox(1)',
      'file:///etc/passwd',
      'not a url',
    ]) {
      expect(hrefs(frame([at({ externalTarget })]))).toEqual([[null, null], [null, null]]);
    }
  });

  it('a refused external target with a location follows the location', () => {
    expect(hrefs(frame([at({ externalTarget: 'javascript:alert(1)', location: 'Sheet1!B2' })]))[0][0]).toBe('#Sheet1!B2');
  });

  it('a range link covers its visible cells and the first link wins', () => {
    const list = frame([
      { top: 0, left: 0, bottom: 0, right: 0, externalTarget: 'https://example.com/first' },
      { top: 0, left: 0, bottom: 9, right: 9, externalTarget: 'https://example.com/range' },
    ]);
    expect(hrefs(list)).toEqual([
      ['https://example.com/first', 'https://example.com/range'],
      ['https://example.com/range', 'https://example.com/range'],
    ]);
  });

  it('follows the frame window, including hidden rows and columns', () => {
    const list = frame([at({ externalTarget: 'https://example.com/far' }, 7, 12)], {
      startRow: 5, startCol: 10, rowIndices: [5, 7], colIndices: [10, 12],
    });
    expect(hrefs(list)).toEqual([[null, null], [null, 'https://example.com/far']]);
  });
});
