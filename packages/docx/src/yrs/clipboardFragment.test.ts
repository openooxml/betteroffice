import { describe, expect, test } from 'bun:test';
import {
  applyClipboardFragment,
  readClipboardSlice,
  readTableFragment,
  validClipboardFragment,
  type ClipboardPasteEngine,
  type ClipboardTableReader,
} from './clipboardFragment';
import type { YrsLoc, YrsStorySegment } from './index';

function textSegment(text: string, attributes: Record<string, unknown> = {}): YrsStorySegment {
  return { kind: 'text', text, attributes };
}

function pilcrowSegment(
  paraId: string,
  properties: Record<string, unknown> = {},
  attributes: Record<string, unknown> = {},
): YrsStorySegment {
  return { kind: 'pilcrow', paraId, properties, attributes };
}

function embedSegment(
  embedKind: string,
  payload: Record<string, unknown> = {},
  attributes: Record<string, unknown> = {},
): YrsStorySegment {
  return { kind: 'embed', embedKind, payload, attributes };
}

function fakeReader(
  stories: Record<string, YrsStorySegment[]>,
  tables: Record<string, Record<string, unknown>>,
): ClipboardTableReader {
  return {
    storySegments: (story) => stories[story] ?? [],
    tablePayload: (story, index) => tables[`${story}:${index}`] ?? null,
  };
}

function payload2x2(): Record<string, unknown> {
  return {
    tblPr: { width: 4000, widthType: 'dxa' },
    grid: [2000, 2000],
    rows: [
      {
        trPr: { height: 10 },
        cells: [
          { story: 'c00', tcPr: { rowspan: 1, colspan: 1 } },
          { story: 'c01', tcPr: { rowspan: 1, colspan: 1 } },
        ],
      },
      {
        trPr: {},
        cells: [
          { story: 'c10', tcPr: { rowspan: 1, colspan: 1 } },
          { story: 'c11', tcPr: { rowspan: 1, colspan: 1 } },
        ],
      },
    ],
  };
}

function stories2x2(): Record<string, YrsStorySegment[]> {
  return {
    c00: [textSegment('a', { bold: true }), pilcrowSegment('p00', { alignment: 'center' })],
    c01: [textSegment('b'), pilcrowSegment('p01')],
    c10: [textSegment('c'), pilcrowSegment('p10')],
    c11: [textSegment('d'), pilcrowSegment('p11')],
  };
}

describe('readTableFragment', () => {
  test('reads a whole table with texts, run attributes and pilcrow properties', () => {
    const reader = fakeReader(stories2x2(), { 'body:0': payload2x2() });
    const fragment = readTableFragment(reader, { story: 'body', tableIndex: 0 });
    expect(fragment.tblPr).toEqual({ width: 4000, widthType: 'dxa' });
    expect(fragment.grid).toEqual([2000, 2000]);
    expect(fragment.rows).toHaveLength(2);
    expect(fragment.rows[0]!.trPr).toEqual({ height: 10 });
    expect(fragment.rows[0]!.cells).toHaveLength(2);
    expect(fragment.rows[0]!.cells[0]).toEqual({
      tcPr: { rowspan: 1, colspan: 1 },
      content: [
        { kind: 'text', text: 'a', attributes: { bold: true } },
        { kind: 'pilcrow', properties: { alignment: 'center' } },
      ],
    });
    expect(fragment.rows[1]!.cells[1]!.content).toEqual([
      { kind: 'text', text: 'd' },
      { kind: 'pilcrow' },
    ]);
    expect(validClipboardFragment({ paragraphs: [{ tables: [fragment], text: '' }] })).toBe(true);
  });

  test('expands a rect to cover a merged cell and narrows the width when columns are cut', () => {
    const payload = {
      tblPr: { width: 6000, widthType: 'dxa' },
      grid: [1000, 2000, 3000],
      rows: [
        { trPr: {}, cells: [{ story: 'm00', tcPr: { rowspan: 1, colspan: 2 } }, { story: 'm02', tcPr: {} }] },
        { trPr: {}, cells: [{ story: 'm10', tcPr: {} }, { story: 'm11', tcPr: {} }, { story: 'm12', tcPr: {} }] },
      ],
    };
    const reader = fakeReader(
      {
        m00: [textSegment('merged'), pilcrowSegment('pm')],
        m02: [textSegment('x'), pilcrowSegment('px')],
        m10: [textSegment('a'), pilcrowSegment('p10')],
        m11: [textSegment('b'), pilcrowSegment('p11')],
        m12: [textSegment('c'), pilcrowSegment('p12')],
      },
      { 'body:0': payload },
    );
    const expanded = readTableFragment(reader, { story: 'body', tableIndex: 0 }, { top: 0, bottom: 0, left: 1, right: 1 });
    expect(expanded.rows).toHaveLength(1);
    expect(expanded.rows[0]!.cells).toHaveLength(1);
    expect(expanded.rows[0]!.cells[0]!.tcPr).toMatchObject({ rowspan: 1, colspan: 2 });
    expect(expanded.grid).toEqual([1000, 2000]);
    expect(expanded.tblPr).toEqual({ width: 3000, widthType: 'dxa' });
    const cut = readTableFragment(reader, { story: 'body', tableIndex: 0 }, { top: 1, bottom: 1, left: 1, right: 1 });
    expect(cut.grid).toEqual([2000]);
    expect(cut.tblPr).toEqual({ width: 2000, widthType: 'dxa' });
    expect(cut.rows[0]!.cells.map((cell) => cell.content[0])).toEqual([{ kind: 'text', text: 'b' }]);
  });

  test('reads nested tables recursively and skips tracked deletions', () => {
    const reader = fakeReader(
      {
        outer: [
          embedSegment('table', {
            tblPr: {},
            grid: [100],
            rows: [{ trPr: {}, cells: [{ story: 'inner', tcPr: {} }] }],
          }),
          pilcrowSegment('pouter'),
        ],
        inner: [
          textSegment('gone', { del: { id: 'r1' } }),
          textSegment('kept'),
          embedSegment('field', { displayText: 'F' }),
          pilcrowSegment('pinner'),
        ],
      },
      {
        'body:0': {
          tblPr: {},
          grid: [100],
          rows: [{ trPr: {}, cells: [{ story: 'outer', tcPr: {} }] }],
        },
      },
    );
    const fragment = readTableFragment(reader, { story: 'body', tableIndex: 0 });
    const nested = fragment.rows[0]!.cells[0]!.content[0];
    expect(nested?.kind).toBe('table');
    if (nested?.kind !== 'table') throw new Error('expected a nested table');
    expect(nested.table.rows[0]!.cells[0]!.content).toEqual([
      { kind: 'text', text: 'kept' },
      { kind: 'text', text: 'F' },
      { kind: 'pilcrow' },
    ]);
  });

  test('extracts cell shading from the background color', () => {
    const reader = fakeReader(
      { shaded: [textSegment('s'), pilcrowSegment('ps')] },
      {
        'body:0': {
          tblPr: {},
          grid: [100],
          rows: [{ trPr: {}, cells: [{ story: 'shaded', tcPr: { backgroundColor: 'FF0000' } }] }],
        },
      },
    );
    const fragment = readTableFragment(reader, { story: 'body', tableIndex: 0 });
    expect(fragment.rows[0]!.cells[0]).toMatchObject({
      tcPr: { rowspan: 1, colspan: 1 },
      shading: 'FF0000',
    });
    expect('backgroundColor' in fragment.rows[0]!.cells[0]!.tcPr).toBe(false);
  });
});

describe('readClipboardSlice', () => {
  test('reads a multi-cell selection as tab-separated rows plus a one-paragraph fragment', () => {
    const reader = fakeReader(stories2x2(), { 'body:0': payload2x2() });
    const slice = readClipboardSlice(reader, {
      cells: {
        anchor: { story: 'body', tableIndex: 0, row: 0, column: 0 },
        head: { story: 'body', tableIndex: 0, row: 1, column: 1 },
      },
    });
    expect(slice.text).toBe('a\tb\nc\td');
    expect(slice.fragment?.paragraphs).toHaveLength(1);
    expect(slice.fragment?.paragraphs[0]).toMatchObject({ text: '' });
    expect(slice.fragment?.paragraphs[0]?.tables).toHaveLength(1);
    expect(slice.fragment?.paragraphs[0]?.tables?.[0]?.rows).toHaveLength(2);
    expect(validClipboardFragment(slice.fragment)).toBe(true);
  });

  test('reads a body range spanning a table paragraph by paragraph', () => {
    const table = payload2x2();
    const reader = fakeReader(
      {
        ...stories2x2(),
        body: [
          textSegment('Hi'),
          pilcrowSegment('p0'),
          embedSegment('table', table),
          textSegment('end'),
          pilcrowSegment('p1'),
        ],
      },
      { 'body:0': table },
    );
    const slice = readClipboardSlice(reader, { story: 'body', from: 0, to: 8 });
    expect(slice.text).toBe('Hi\n' + 'a\tb\nc\td\n' + 'end\n');
    expect(slice.fragment?.paragraphs).toEqual([
      { text: 'Hi' },
      { tables: expect.any(Array), text: 'end' },
    ]);
    expect(slice.fragment?.paragraphs[1]?.tables?.[0]?.rows).toHaveLength(2);
  });

  test('returns no fragment when the range holds no table', () => {
    const reader = fakeReader({ body: [textSegment('Hi'), pilcrowSegment('p0')] }, {});
    const slice = readClipboardSlice(reader, { story: 'body', from: 0, to: 3 });
    expect(slice).toEqual({ text: 'Hi\n', fragment: null });
  });

  test('maps tabs, breaks and other embeds to paragraph structure', () => {
    const reader = fakeReader(
      {
        body: [
          textSegment('a'),
          embedSegment('tab'),
          embedSegment('break'),
          embedSegment('math', { plainText: 'x\ny' }),
          embedSegment('table', payload2x2()),
          pilcrowSegment('p0'),
        ],
      },
      { 'body:0': payload2x2() },
    );
    const slice = readClipboardSlice(reader, { story: 'body', from: 0, to: 7 });
    expect(slice.fragment?.paragraphs).toEqual([
      { text: 'a\t' },
      { text: 'x' },
      { text: 'y' },
      { tables: expect.any(Array), text: '' },
    ]);
  });
});

describe('vertically merged rows on copy', () => {
  test('keeps a fully covered row and the rowspan of its covering cell', () => {
    const reader = fakeReader(
      { a: [textSegment('A'), pilcrowSegment('pa')] },
      {
        'body:0': {
          tblPr: {},
          grid: [2000],
          rows: [
            { trPr: { height: 10 }, cells: [{ story: 'a', tcPr: { rowspan: 2, colspan: 1 } }] },
            { trPr: { height: 20 }, cells: [] },
          ],
        },
      },
    );
    const fragment = readTableFragment(reader, { story: 'body', tableIndex: 0 });
    expect(fragment.rows).toHaveLength(2);
    expect(fragment.rows[0]!.cells[0]!.tcPr).toMatchObject({ rowspan: 2 });
    expect(fragment.rows[1]).toEqual({ trPr: { height: 20 }, cells: [] });
    expect(validClipboardFragment({ paragraphs: [{ tables: [fragment], text: '' }] })).toBe(true);
  });

  test('keeps the covered row of a 2x2 table with a merged first column', () => {
    const reader = fakeReader(
      {
        a: [textSegment('A'), pilcrowSegment('pa')],
        b: [textSegment('B'), pilcrowSegment('pb')],
        c: [textSegment('C'), pilcrowSegment('pc')],
      },
      {
        'body:0': {
          tblPr: {},
          grid: [2000, 2000],
          rows: [
            {
              trPr: {},
              cells: [
                { story: 'a', tcPr: { rowspan: 2, colspan: 1 } },
                { story: 'b', tcPr: { rowspan: 1, colspan: 1 } },
              ],
            },
            { trPr: {}, cells: [{ story: 'c', tcPr: { rowspan: 1, colspan: 1 } }] },
          ],
        },
      },
    );
    const fragment = readTableFragment(reader, { story: 'body', tableIndex: 0 });
    expect(fragment.rows).toHaveLength(2);
    expect(fragment.rows[0]!.cells[0]!.tcPr).toMatchObject({ rowspan: 2 });
    expect(fragment.rows[1]!.cells).toHaveLength(1);
    expect(fragment.rows[1]!.cells[0]!.content[0]).toEqual({ kind: 'text', text: 'C' });
    expect(validClipboardFragment({ paragraphs: [{ tables: [fragment], text: '' }] })).toBe(true);
  });
});

describe('validClipboardFragment spans', () => {
  function fragmentWithSpan(tcPr: Record<string, unknown>): unknown {
    return {
      paragraphs: [{
        tables: [{
          tblPr: {},
          grid: [],
          rows: [{ trPr: {}, cells: [{ tcPr, content: [{ kind: 'text', text: 'a' }, { kind: 'pilcrow' }] }] }],
        }],
        text: '',
      }],
    };
  }

  test('rejects spans outside 1..=1000', () => {
    expect(validClipboardFragment(fragmentWithSpan({ colspan: 1_000_000_000 }))).toBe(false);
    expect(validClipboardFragment(fragmentWithSpan({ rowspan: 1_000_000_000 }))).toBe(false);
    expect(validClipboardFragment(fragmentWithSpan({ colspan: 0 }))).toBe(false);
    expect(validClipboardFragment(fragmentWithSpan({ rowspan: 1.5 }))).toBe(false);
    expect(validClipboardFragment(fragmentWithSpan({ colspan: '2' }))).toBe(false);
    expect(validClipboardFragment(fragmentWithSpan({ colspan: 1000 }))).toBe(true);
    expect(validClipboardFragment(fragmentWithSpan({ rowspan: 1, colspan: 1 }))).toBe(true);
    expect(validClipboardFragment(fragmentWithSpan({}))).toBe(true);
  });
});

describe('validClipboardFragment rust parity', () => {
  function tableFragment(rows: unknown, grid: unknown = []): unknown {
    return { paragraphs: [{ tables: [{ tblPr: {}, grid, rows }], text: '' }] };
  }
  function cell(tcPr: Record<string, unknown>, format?: Record<string, unknown>): unknown {
    return {
      tcPr,
      content: [
        format === undefined
          ? { kind: 'text', text: 'a' }
          : { kind: 'text', text: 'a', format },
        { kind: 'pilcrow' },
      ],
    };
  }

  test('rejects a rowspan past the last row', () => {
    expect(validClipboardFragment(tableFragment([
      { trPr: {}, cells: [cell({ rowspan: 2 })] },
    ]))).toBe(false);
  });

  test('rejects overlapping anchors', () => {
    expect(validClipboardFragment(tableFragment([
      { trPr: {}, cells: [cell({ colspan: 2 }), cell({ rowspan: 2 })] },
      { trPr: {}, cells: [cell({ colspan: 3 })] },
    ]))).toBe(false);
  });

  test('rejects an uncovered empty row and a table without cells', () => {
    expect(validClipboardFragment(tableFragment([
      { trPr: {}, cells: [cell({})] },
      { trPr: {}, cells: [] },
    ]))).toBe(false);
    expect(validClipboardFragment(tableFragment([
      { trPr: {}, cells: [] },
    ]))).toBe(false);
  });

  test('rejects a grid whose length differs from the column count', () => {
    const rows = [{ trPr: {}, cells: [cell({}), cell({})] }];
    expect(validClipboardFragment(tableFragment(rows, [2000]))).toBe(false);
    expect(validClipboardFragment(tableFragment(rows, [2000, 2000]))).toBe(true);
    expect(validClipboardFragment(tableFragment(rows, []))).toBe(true);
  });

  test('rejects mistyped format delta values', () => {
    const withFormat = (format: Record<string, unknown>): unknown =>
      tableFragment([{ trPr: {}, cells: [cell({}, format)] }]);
    expect(validClipboardFragment(withFormat({ bold: 'yes' }))).toBe(false);
    expect(validClipboardFragment(withFormat({ fontSize: -1 }))).toBe(false);
    expect(validClipboardFragment(withFormat({ fontSize: '12' }))).toBe(false);
    expect(validClipboardFragment(withFormat({ color: { rgb: 'FF0000', themeColor: 'accent1' } }))).toBe(false);
    expect(validClipboardFragment(withFormat({ color: {} }))).toBe(false);
    expect(validClipboardFragment(withFormat({ fontFamily: {} }))).toBe(false);
    expect(validClipboardFragment(withFormat({ underline: { style: 3 } }))).toBe(false);
    expect(validClipboardFragment(withFormat({ bold: true, fontSize: 12 }))).toBe(true);
    expect(validClipboardFragment(withFormat({ color: { rgb: 'FF0000' } }))).toBe(true);
    expect(validClipboardFragment(withFormat({ fontFamily: { ascii: 'Calibri' } }))).toBe(true);
  });

  test('an invalid second table aborts the paste before any mutation', () => {
    const calls: string[] = [];
    const good = {
      tblPr: {}, grid: [],
      rows: [{ trPr: {}, cells: [{ tcPr: {}, content: [{ kind: 'text', text: 'a' }, { kind: 'pilcrow' }] }] }],
    };
    const bad = {
      tblPr: {}, grid: [],
      rows: [{ trPr: {}, cells: [{ tcPr: { rowspan: 2 }, content: [{ kind: 'text', text: 'b' }, { kind: 'pilcrow' }] }] }],
    };
    const engine = {
      validateTableFragment: (table: unknown) => {
        calls.push('validate');
        if (table === bad) throw new Error('invalid table');
      },
      deleteRange: () => {
        calls.push('delete');
        return { revisionId: null, range: { story: 'body', start: { paraId: 'p0', offset: 0 }, end: { paraId: 'p0', offset: 0 } } };
      },
      splitParagraph: (at: YrsLoc) => ({ firstParaId: at.paraId, secondParaId: 'p1', revisionId: null }),
      insertText: (at: YrsLoc, text: string) => ({
        revisionId: null,
        range: { story: at.story, start: at, end: { paraId: at.paraId, offset: at.offset + text.length } },
      }),
      insertTableFragment: (at: YrsLoc) => {
        calls.push('table');
        return {
          table: { story: at.story, tableIndex: 0 }, rows: 1, columns: 1,
          createdStoryIds: [], deletedStoryIds: [], changedStoryIds: [], newParaIds: [],
          deletedTable: false, revisionIds: [],
          after: { story: at.story, paraId: at.paraId, offset: at.offset + 1 },
        };
      },
    } as unknown as ClipboardPasteEngine;
    expect(() => applyClipboardFragment(
      engine,
      { story: 'body', start: { paraId: 'p0', offset: 0 }, end: { paraId: 'p0', offset: 2 } },
      { paragraphs: [{ tables: [good as never], text: '' }, { tables: [bad as never], text: '' }] },
    )).toThrow('invalid table');
    expect(calls).toEqual(['validate', 'validate']);
  });
});

describe('applyClipboardFragment', () => {
  function fakeEngine(calls: string[]): ClipboardPasteEngine {
    const engine = {
      validateTableFragment: () => {},
      deleteRange: (range: { story: string; start: { paraId: string; offset: number } }) => {
        calls.push('delete');
        return {
          revisionId: null,
          range: { story: range.story, start: range.start, end: range.start },
        };
      },
      splitParagraph: (at: YrsLoc) => {
        calls.push('split');
        return { firstParaId: at.paraId, secondParaId: `${at.paraId}n`, revisionId: null };
      },
      insertText: (at: YrsLoc, text: string) => {
        calls.push(`text:${text}`);
        return {
          revisionId: null,
          range: {
            story: at.story,
            start: { paraId: at.paraId, offset: at.offset },
            end: { paraId: at.paraId, offset: at.offset + text.length },
          },
        };
      },
      insertTableFragment: (at: YrsLoc) => {
        calls.push('table');
        return {
          table: { story: at.story, tableIndex: 0 },
          rows: 1,
          columns: 1,
          createdStoryIds: [],
          deletedStoryIds: [],
          changedStoryIds: [],
          newParaIds: [],
          deletedTable: false,
          revisionIds: [],
          after: { story: at.story, paraId: at.paraId, offset: at.offset + 1 },
        };
      },
    };
    return engine as unknown as ClipboardPasteEngine;
  }

  test('deletes an expanded range, then replays tables, text and splits', () => {
    const calls: string[] = [];
    const caret = applyClipboardFragment(
      fakeEngine(calls),
      {
        story: 'body',
        start: { paraId: 'p0', offset: 1 },
        end: { paraId: 'p0', offset: 3 },
      },
      {
        paragraphs: [
          { tables: [payload2x2() as never], text: 'a' },
          { text: 'b' },
        ],
      },
    );
    expect(calls).toEqual(['delete', 'table', 'text:a', 'split', 'text:b']);
    expect(caret).toEqual({ story: 'body', paraId: 'p0n', offset: 1 });
  });

  test('leaves a collapsed range alone and skips empty text', () => {
    const calls: string[] = [];
    const caret = applyClipboardFragment(
      fakeEngine(calls),
      { story: 'body', start: { paraId: 'p0', offset: 2 }, end: { paraId: 'p0', offset: 2 } },
      { paragraphs: [{ tables: [payload2x2() as never], text: '' }] },
    );
    expect(calls).toEqual(['table']);
    expect(caret).toEqual({ story: 'body', paraId: 'p0', offset: 3 });
  });
});
