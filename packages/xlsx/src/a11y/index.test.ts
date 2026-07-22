import { describe, expect, it } from 'bun:test';
import { buildA11yGrid } from './index';
import type { A11yStrings } from './index';
import type { DisplayList, DrawCmd, Rect } from '../display-list/types';
import type { Selection } from '../selection/index';

const strings: A11yStrings = {
  gridLabel: 'Spreadsheet grid',
  cellLabel: '{address}, {value}',
  cellLabelSelected: '{address}, {value}, selected',
  emptyCellLabel: '{address}, empty',
  emptyCellLabelSelected: '{address}, empty, selected',
  rowHeaderLabel: 'Row {row}',
  columnHeaderLabel: 'Column {column}',
};

function textCmd(clip: Rect, value: string): DrawCmd {
  return {
    op: 'text',
    x: clip.x + 2,
    y: clip.y + 12,
    text: value,
    fontSize: 12,
    color: '#000',
    clip,
  };
}

// a 2x2 window at sheet origin (0,0): B1 and A2 carry text, the rest are empty.
function sampleDisplayList(): DisplayList {
  const box = (c: number, r: number): Rect => ({ x: c * 80, y: r * 20, w: 80, h: 20 });
  return {
    width: 160,
    height: 40,
    commands: [textCmd(box(1, 0), '42'), textCmd(box(0, 1), 'hi')],
    grid: { startRow: 0, startCol: 0, colOffsets: [0, 80, 160], rowOffsets: [0, 20, 40] },
  };
}

describe('buildA11yGrid', () => {
  it('mirrors the grid dimensions, headers, and addresses', () => {
    const g = buildA11yGrid(sampleDisplayList(), null, 'Sheet1', strings);
    expect(g.sheetName).toBe('Sheet1');
    expect(g.label).toBe('Spreadsheet grid');
    expect(g.columnHeaders.map((h) => h.text)).toEqual(['A', 'B']);
    expect(g.columnHeaders[1].label).toBe('Column B');
    expect(g.rows.map((r) => r.header)).toEqual(['Row 1', 'Row 2']);
    expect(g.rows[0].cells.map((c) => c.address)).toEqual(['A1', 'B1']);
    expect(g.rows[1].cells.map((c) => c.address)).toEqual(['A2', 'B2']);
  });

  it('recovers cell text from the clipped text commands', () => {
    const g = buildA11yGrid(sampleDisplayList(), null, 'Sheet1', strings);
    expect(g.rows[0].cells[1].text).toBe('42');
    expect(g.rows[1].cells[0].text).toBe('hi');
    expect(g.rows[0].cells[0].text).toBe('');
  });

  it('labels filled and empty cells distinctly', () => {
    const g = buildA11yGrid(sampleDisplayList(), null, 'Sheet1', strings);
    expect(g.rows[0].cells[1].label).toBe('B1, 42');
    expect(g.rows[0].cells[0].label).toBe('A1, empty');
  });

  it('marks the selected range and appends "selected" to its labels', () => {
    const selection: Selection = { anchor: { row: 0, col: 1 }, focus: { row: 1, col: 1 } };
    const g = buildA11yGrid(sampleDisplayList(), selection, 'Sheet1', strings);
    expect(g.rows[0].cells[1].selected).toBe(true);
    expect(g.rows[0].cells[1].label).toBe('B1, 42, selected');
    expect(g.rows[1].cells[1].selected).toBe(true);
    expect(g.rows[1].cells[1].label).toBe('B2, empty, selected');
    expect(g.rows[0].cells[0].selected).toBe(false);
  });

  it('returns an empty grid when the frame has no grid metadata', () => {
    const g = buildA11yGrid({ width: 0, height: 0, commands: [] }, null, 'Sheet1', strings);
    expect(g.rows).toEqual([]);
    expect(g.columnHeaders).toEqual([]);
    expect(g.label).toBe('Spreadsheet grid');
  });

  it('skips ghost preview commands so the committed text wins', () => {
    const dl = sampleDisplayList();
    const box: Rect = { x: 80, y: 0, w: 80, h: 20 };
    const preview = textCmd(box, '99');
    if (preview.op === 'text') preview.ghost = true;
    dl.commands.push(preview);
    const g = buildA11yGrid(dl, null, 'Sheet1', strings);
    expect(g.rows[0].cells[1].text).toBe('42');
  });

  it('addresses columns past Z with bijective letters', () => {
    const dl: DisplayList = {
      width: 80,
      height: 20,
      commands: [],
      grid: { startRow: 0, startCol: 26, colOffsets: [0, 80], rowOffsets: [0, 20] },
    };
    const g = buildA11yGrid(dl, null, 'S', strings);
    expect(g.columnHeaders[0].text).toBe('AA');
    expect(g.rows[0].cells[0].address).toBe('AA1');
  });
});
