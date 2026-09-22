import { expect } from 'bun:test';

import { STORY, UNDO_STEP_GAP_MS, displayFrame, frameText, regionLayout, typingTarget } from './context';
import type { DocxScenario } from './context';

const ROWS = 5;
const COLUMNS = 4;
const CELLS = 6;

export const tablesDeep: DocxScenario = {
  name: 'tables-deep',
  description: 'Inserts a five by four table, types into six of its cells, grows it by a row and a column, merges a block, shades it, then undoes each structural step and checks the story count and cell text after every one.',
  participants: ['web'],
  async run({ recorder, open }) {
    const editor = await recorder.loadAsync(() => open());
    const { session } = editor;
    const target = typingTarget(session, session.paragraphs(STORY), session.paragraphSpans(STORY));
    const storiesBefore = session.storyIds().length;

    const table = recorder.op('insertTable', () => session.insertTable({ story: STORY, paraId: target.paraId, offset: 0 }, ROWS, COLUMNS));
    expect(table).toMatchObject({ rows: ROWS, columns: COLUMNS, deletedTable: false });
    expect(table.createdStoryIds).toHaveLength(ROWS * COLUMNS);
    expect(session.storyIds()).toHaveLength(storiesBefore + ROWS * COLUMNS);

    const cells = table.createdStoryIds.slice(0, CELLS);
    cells.forEach((storyId, index) => {
      const paraId = session.paragraphs(storyId)[0].paraId;
      recorder.op('insertText:cell', () => session.insertText({ story: storyId, paraId, offset: 0 }, `cell ${index}`));
      expect(session.paragraphs(storyId)[0].text).toBe(`cell ${index}`);
    });

    const tableLoc = table.table;
    const topLeft = { ...tableLoc, row: 0, column: 0 };
    await Bun.sleep(UNDO_STEP_GAP_MS);
    const grown = recorder.op('insertRow', () => session.insertRow(topLeft, 'below'));
    expect(grown.rows).toBe(ROWS + 1);
    expect(session.storyIds()).toHaveLength(storiesBefore + (ROWS + 1) * COLUMNS);
    await Bun.sleep(UNDO_STEP_GAP_MS);
    const wider = recorder.op('insertColumn', () => session.insertColumn(topLeft, 'right'));
    expect(wider.columns).toBe(COLUMNS + 1);
    const widened = session.storyIds().length;
    expect(widened).toBe(storiesBefore + (ROWS + 1) * (COLUMNS + 1));

    await Bun.sleep(UNDO_STEP_GAP_MS);
    const block = { anchor: { ...tableLoc, row: 1, column: 1 }, head: { ...tableLoc, row: 2, column: 2 } };
    const merged = recorder.op('mergeCells', () => session.mergeCells(block));
    expect(merged.table).toEqual(tableLoc);
    recorder.op('setCellShading', () => session.setCellShading(block, '#ffee88'));
    recorder.op('setColumnWidth', () => session.setColumnWidth(topLeft, 2400));
    recorder.op('setTableWidth', () => session.setTableWidth(tableLoc, 9000));

    const laid = regionLayout(editor, recorder, 'layoutDocumentWithRegions:withTable');
    expect(laid.layout.pages.length).toBeGreaterThan(0);
    const frame = displayFrame(session, recorder, 'displayListFrame:withTable', null);
    expect(frameText(frame)).toContain('cell 0');

    await Bun.sleep(UNDO_STEP_GAP_MS);
    expect(recorder.op('undo:formatting', () => session.undo())).toBe(true);
    expect(session.storyIds()).toHaveLength(widened);
    expect(recorder.op('undo:insertColumn', () => session.undo())).toBe(true);
    expect(session.storyIds()).toHaveLength(storiesBefore + (ROWS + 1) * COLUMNS);
    expect(recorder.op('undo:insertRow', () => session.undo())).toBe(true);
    expect(session.storyIds()).toHaveLength(storiesBefore + ROWS * COLUMNS);
    expect(session.paragraphs(cells[0])[0].text).toBe('cell 0');

    const removed = recorder.op('deleteTable', () => session.deleteTable(tableLoc));
    expect(removed.deletedTable).toBe(true);
    expect(session.storyIds()).toHaveLength(storiesBefore);
    regionLayout(editor, recorder, 'layoutDocumentWithRegions:afterDelete');
    expect(frameText(displayFrame(session, recorder, 'displayListFrame:afterDelete', frame))).not.toContain('cell 0');
  },
};
