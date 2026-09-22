import { expect } from 'bun:test';

import { PRINT_METRICS, SHEET, profiledDisplayList } from './context';
import type { XlsxScenario } from './context';

const STEPS = 30;

export const viewportScrollRender: XlsxScenario = {
  name: 'viewport-scroll-render',
  description: 'Builds thirty display lists while scrolling down and right, then a wide viewport and a print range, checking that the grid metadata advances monotonically.',
  participants: ['web'],
  run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const info = recorder.op('sheetInfo', () => handle.sheetInfo());
    let previousRow = -1;
    let previousCol = -1;
    let firstRow = -1;
    for (let step = 0; step < STEPS; step += 1) {
      const viewport = { x: step * 120, y: step * 260, width: 1280, height: 800 };
      const list = profiledDisplayList(handle, recorder, 'displayList:scroll', viewport);
      expect(list.grid).toBeDefined();
      const grid = list.grid!;
      expect(grid.rowOffsets.length).toBe((grid.rowIndices?.length ?? grid.rowOffsets.length - 1) + 1);
      const lastRow = grid.rowIndices?.at(-1) ?? grid.startRow + grid.rowOffsets.length - 2;
      const lastCol = grid.colIndices?.at(-1) ?? grid.startCol + grid.colOffsets.length - 2;
      expect(lastRow).toBeGreaterThanOrEqual(previousRow);
      expect(lastCol).toBeGreaterThanOrEqual(previousCol);
      expect(grid.rowOffsets.every((offset, index) => index === 0 || offset >= grid.rowOffsets[index - 1])).toBe(true);
      if (firstRow < 0) firstRow = lastRow;
      previousRow = lastRow;
      previousCol = lastCol;
    }
    expect(previousRow).toBeGreaterThan(firstRow);

    const wide = profiledDisplayList(handle, recorder, 'displayList:wide', { x: 0, y: 0, width: 2560, height: 1440 });
    expect(wide.commands.length).toBeGreaterThan(0);
    expect(wide.width).toBeGreaterThan(0);

    const printed = recorder.op('printDisplayList', () => handle.printDisplayList(SHEET, 'A1:H40', PRINT_METRICS, true));
    expect(printed.commands.length).toBeGreaterThan(0);

    const charts = wide.charts ?? [];
    if (charts.length > 0) {
      const chart = charts[0];
      const hit = recorder.op('chartAtPoint', () => handle.chartAtPoint({ x: 0, y: 0, width: 2560, height: 1440 }, chart.rect.x + 4, chart.rect.y + 4));
      expect(hit?.id).toBe(chart.id);
      expect(recorder.op('moveChart', () => handle.moveChart(SHEET, chart.id, 24, 24)).applied).toBe(true);
      expect(recorder.op('undo:moveChart', () => handle.undo()).applied).toBe(true);
    }
    expect(info.sheetNames.length).toBeGreaterThan(0);
  },
};
