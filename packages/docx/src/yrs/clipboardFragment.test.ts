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
      { text: '' },
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
      { text: '' },
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


describe('picture clipboard fragments', () => {
  const src = 'data:image/png;base64,AA==';
  test('picture-only copy resolves media and removes drawing identity', () => {
    const reader = { ...fakeReader({ body: [embedSegment('image', { src: 'media:0', rId: 'rId1', _kind: 'image', id: 'old', embedId: 'old', docPrId: 1, drawingId: 1, width: 120, height: 60 }), pilcrowSegment('p0')] }, {}),
      mediaDataUrl: () => src };
    const slice = readClipboardSlice(reader, { story: 'body', from: 0, to: 1 });
    expect(slice.text).toBe('');
    expect(slice.fragment).toEqual({ paragraphs: [{ text: '', inlines: [{ image: { src, width: 120, height: 60 } }] }] });
    expect(validClipboardFragment(slice.fragment)).toBe(true);
  });

  test('mixed copy preserves text and picture order', () => {
    const reader = fakeReader({ body: [textSegment('ab'), embedSegment('image', { src }), textSegment('cd'), pilcrowSegment('p0')] }, {});
    expect(readClipboardSlice(reader, { story: 'body', from: 0, to: 5 })).toEqual({
      text: 'abcd', fragment: { paragraphs: [{ text: 'abcd', inlines: [{ text: 'ab' }, { image: { src } }, { text: 'cd' }] }] },
    });
  });

  test('unresolved picture copy fails instead of dropping media', () => {
    const reader = fakeReader({ body: [embedSegment('image', { src: 'media:9' })] }, {});
    expect(() => readClipboardSlice(reader, { story: 'body', from: 0, to: 1 })).toThrow();
  });

  test('picture validation rejects unsafe sources and oversized media', () => {
    const fragment = (image: Record<string, unknown>) => ({ paragraphs: [{ text: '', inlines: [{ image }] }] });
    for (const source of ['https://invalid/image.png', 'http://invalid/image.png', 'blob:picture', 'media:0', 'data:text/html;base64,AA==']) {
      expect(validClipboardFragment(fragment({ src: source }))).toBe(false);
    }
    expect(validClipboardFragment(fragment({ src, _kind: 'image' }))).toBe(false);
    expect(validClipboardFragment(fragment({ src, widthEmu: -1 }))).toBe(false);
    expect(validClipboardFragment(fragment({ src: 'data:image/png;base64,' + 'A'.repeat(64 * 1024 * 1024) }))).toBe(false);
    expect(validClipboardFragment(fragment({ src }))).toBe(true);
  });
});


test('oversized table copies retain their text and fragment without pictures', () => {
  const text = 'a'.repeat(1_000_001);
  const table = { tblPr: {}, grid: [], rows: [{ trPr: {}, cells: [{ story: 'cell', tcPr: {} }] }] };
  const reader = fakeReader({
    body: [embedSegment('table', table), embedSegment('image', { src: 'data:image/png;base64,AA==' })],
    cell: [textSegment(text), pilcrowSegment('pc')],
  }, { 'body:0': table });
  const slice = readClipboardSlice(reader, { story: 'body', from: 0, to: 1 });
  expect(slice.text).toContain(text);
  expect(slice.fragment?.paragraphs[0]?.tables?.[0]?.rows[0]?.cells[0]?.content[0]).toEqual({ kind: 'text', text });
  expect(validClipboardFragment(slice.fragment)).toBe(false);
  expect(() => readClipboardSlice(reader, { story: 'body', from: 0, to: 2 })).toThrow();
});


test('picture paragraph copy includes the selected trailing paragraph mark', () => {
  const image = { src: 'data:image/png;base64,AA==' };
  const reader = fakeReader({ body: [embedSegment('image', image), pilcrowSegment('p0'), textSegment('next'), pilcrowSegment('p1')] }, {});
  expect(readClipboardSlice(reader, { story: 'body', from: 0, to: 2 })).toEqual({
    text: '\n', fragment: { paragraphs: [{ text: '', inlines: [{ image }] }, { text: '' }] },
  });
  expect(readClipboardSlice(reader, { story: 'body', from: 0, to: 1 }).fragment?.paragraphs).toHaveLength(1);
});

test('table paragraph copy includes the selected trailing paragraph mark', () => {
  const reader = fakeReader({ ...stories2x2(), body: [embedSegment('table', payload2x2()), pilcrowSegment('p0'), textSegment('next')] }, { 'body:0': payload2x2() });
  expect(readClipboardSlice(reader, { story: 'body', from: 0, to: 2 }).fragment?.paragraphs).toEqual([
    { text: '', tables: [readTableFragment(reader, { story: 'body', tableIndex: 0 })] }, { text: '' },
  ]);
});

test('picture copy preserves consecutive selected breaks and text-only copy remains plain', () => {
  const image = { src: 'data:image/png;base64,AA==' };
  const reader = fakeReader({ body: [embedSegment('image', image), embedSegment('break'), pilcrowSegment('p0')] }, {});
  expect(readClipboardSlice(reader, { story: 'body', from: 0, to: 3 }).fragment?.paragraphs).toEqual([
    { text: '', inlines: [{ image }] }, { text: '' }, { text: '' },
  ]);
  expect(readClipboardSlice(reader, { story: 'body', from: 1, to: 3 })).toEqual({ text: '\n\n', fragment: null });
});

test('picture copy rejects unrepresentable visible embeds in either order', () => {
  for (const kind of ['chart', 'shape', 'opaqueDrawing', 'horizontalRule', 'noteRef', 'pageBreak', 'columnBreak', 'opaque', 'field', 'math', 'sdt', 'blockSdt', 'unknown']) {
    for (const first of [false, true]) {
      const image = embedSegment('image', { src: 'data:image/png;base64,AA==' });
      const other = embedSegment(kind);
      const reader = fakeReader({ body: first ? [image, other] : [other, image] }, {});
      expect(() => readClipboardSlice(reader, { story: 'body', from: 0, to: 2 })).toThrow();
      expect(readClipboardSlice(reader, { story: 'body', from: first ? 1 : 0, to: first ? 2 : 1 })).toEqual({ text: '', fragment: null });
    }
  }
});

for (const objectKind of ['chart', 'shape', 'image']) {
  for (const depth of [0, 1, 3]) {
    for (const controlKind of [null, 'sdt', 'blockSdt']) {
      test(`picture copy rejects ${objectKind} in a table cell at depth ${depth} through ${controlKind ?? 'no control'}`, () => {
        const image = { src: 'data:image/png;base64,AA==' };
        const table = (story: string) => ({ grid: [2000], rows: [{ cells: [{ story, tcPr: {} }] }] });
        const stories: Record<string, YrsStorySegment[]> = {};
        for (let level = 0; level < depth; level++) stories[`cell${level}`] = [embedSegment('table', table(`cell${level + 1}`))];
        const object = embedSegment(objectKind, image);
        stories[`cell${depth}`] = controlKind === 'sdt'
          ? [embedSegment('sdt', { content: [{ kind: 'text', text: 'kept' }, { kind: objectKind, payload: image }] })]
          : controlKind === 'blockSdt'
            ? [embedSegment('blockSdt', { story: 'control' })]
            : [object];
        stories.control = [textSegment('kept'), object];
        for (const first of [false, true]) {
          const picture = embedSegment('image', image);
          const selectedTable = embedSegment('table', table('cell0'));
          stories.body = first ? [picture, selectedTable] : [selectedTable, picture];
          const reader = fakeReader(stories, {});
          expect(() => readClipboardSlice(reader, { story: 'body', from: 0, to: 2 }))
            .toThrow('Picture selection contains unsupported clipboard objects');
          expect(() => readClipboardSlice(reader, { story: 'body', from: first ? 0 : 1, to: first ? 1 : 2 })).not.toThrow();
          expect(() => readClipboardSlice(reader, { story: 'body', from: first ? 1 : 0, to: first ? 2 : 1 })).not.toThrow();
        }
      });
    }
  }
}

test('picture copy permits bookmarks and non-visible comment and range markers', () => {
  for (const kind of ['bookmark', 'bookmarkStart', 'bookmarkEnd', 'commentAnchor', 'commentReference', 'commentRangeStart', 'commentRangeEnd',
    'rangeMarker', 'rangeStart', 'rangeEnd', 'permStart', 'permEnd', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd']) {
    const image = { src: 'data:image/png;base64,AA==' };
    const reader = fakeReader({ body: [embedSegment(kind), embedSegment('image', image)] }, {});
    expect(readClipboardSlice(reader, { story: 'body', from: 0, to: 2 })).toEqual({
      text: '', fragment: { paragraphs: [{ text: '', inlines: [{ image }] }] },
    });
  }
  const reader = fakeReader({ body: [embedSegment('field', { modelKind: 'commentReference', displayText: '' }), embedSegment('image', { src: 'data:image/png;base64,AA==' })] }, {});
  expect(readClipboardSlice(reader, { story: 'body', from: 0, to: 2 }).fragment).not.toBeNull();
});


test('picture copy preserves text-only content controls and nested text controls', () => {
  const image = { src: 'data:image/png;base64,AA==' };
  const reader = fakeReader({
    body: [embedSegment('image', image), embedSegment('sdt', { content: [
      { kind: 'text', text: 'before' }, { kind: 'tab' },
      { kind: 'sdt', payload: { content: [{ kind: 'text', text: 'after' }, { kind: 'break' }, { kind: 'text', text: 'end' }, { kind: 'pilcrow' }] } },
    ] })],
  }, {});
  const slice = readClipboardSlice(reader, { story: 'body', from: 0, to: 2 });
  expect(slice.text).toBe('before\tafter\nend');
  expect(slice.fragment?.paragraphs).toEqual([
    { text: 'before\tafter', inlines: [{ image }, { text: 'before\tafter' }] }, { text: 'end' },
  ]);
  const block = fakeReader({ body: [embedSegment('image', image), embedSegment('blockSdt', { story: 'control' })],
    control: [textSegment('block'), embedSegment('sdt', { content: [{ kind: 'text', text: ' text' }] }), pilcrowSegment('pc')],
  }, {});
  expect(readClipboardSlice(block, { story: 'body', from: 0, to: 2 }).text).toBe('block text\n');
  expect(readClipboardSlice(block, { story: 'body', from: 0, to: 2 }).fragment).not.toBeNull();
});

test('picture copy rejects content controls containing nested pictures and other objects', () => {
  for (const controlKind of ['sdt', 'blockSdt']) {
    for (const objectKind of ['image', 'chart', 'shape', 'table', 'unknown']) {
      for (const nested of [false, true]) {
        const object = { kind: objectKind, payload: { src: 'data:image/png;base64,AA==' } };
        const content = [{ kind: 'text', text: 'kept' }, ...(nested ? [{ kind: 'sdt', payload: { content: [object] } }] : [object])];
        const payload = controlKind === 'sdt' ? { content } : { story: 'control' };
        const child = nested ? embedSegment('sdt', { content: [object] }) : embedSegment(objectKind, object.payload);
        const reader = fakeReader({ body: [embedSegment('image', { src: 'data:image/png;base64,AA==' }), embedSegment(controlKind, payload)],
          control: [textSegment('kept'), child, pilcrowSegment('pc')],
        }, {});
        expect(() => readClipboardSlice(reader, { story: 'body', from: 0, to: 2 })).toThrow();
        const plain = readClipboardSlice(reader, { story: 'body', from: 1, to: 2 });
        expect(plain.text).toContain('kept');
        expect(plain.fragment).toBeNull();
      }
    }
  }
});

test('picture copy rejects content controls beyond the nested text depth limit', () => {
  let payload: Record<string, unknown> = { content: [{ kind: 'text', text: 'deep' }] };
  for (let depth = 0; depth < 10; depth += 1) payload = { content: [{ kind: 'sdt', payload }] };
  const reader = fakeReader({ body: [embedSegment('image', { src: 'data:image/png;base64,AA==' }), embedSegment('sdt', payload)] }, {});
  expect(() => readClipboardSlice(reader, { story: 'body', from: 0, to: 2 })).toThrow();
  expect(readClipboardSlice(reader, { story: 'body', from: 1, to: 2 })).toEqual({ text: 'deep', fragment: null });
});
