import { expect, spyOn, test } from 'bun:test';
import type { PositionedTextLine, SlideDisplayList, TextBoxPrimitive } from '../types';
import * as wasm from '../wasm/loader';
import { buildCommandFrame } from './commandFrame';
import { decodeCommandDisplayList, hitGeometry, indexGeometry, indexHitGeometry, navigateGeometry, selectionGeometry } from './commandGeometry';
import type { PresentationCommandSelection } from './commandProtocol';

const selection: PresentationCommandSelection = { slideId: 'slide', shapeId: 'box', storyId: 'story',
  anchor: 8, focus: 8, affinity: 'downstream' };

function line(start: number, end: number, y: number): PositionedTextLine {
  return { start, end, x: 0, y, width: (end - start) * 10, height: 20, baseline: y + 15,
    runs: [], caretStops: Array.from({ length: end - start + 1 }, (_, index) =>
      ({ position: start + index, x: index * 10 })) };
}

function list(text = 'alpha beta\nx\nalpha beta', lines = [line(0, 10, 0), line(11, 12, 20), line(13, 23, 40)]): SlideDisplayList {
  const box: TextBoxPrimitive = { kind: 'textBox', objectId: 1, shapeId: 'box', storyId: 'story',
    x: 0, y: 0, w: 100, h: 60, anchor: 'top', lines,
    paragraphs: text.split('\n').map((text) => ({ level: 0,
      runs: [{ text, fontFamily: 'Arial', fontSizePt: 12, color: '#000000' }] })) };
  return { contractVersion: 1, width: 960, height: 540,
    background: { kind: 'solid', color: '#ffffff' }, primitives: [box] };
}

test('vertical_goal_worker', () => {
  const index = indexHitGeometry(list());
  const first = navigateGeometry(index, selection, 'down', false)!;
  expect(first.focus).toBe(12);
  expect(navigateGeometry(index, first, 'down', true)).toMatchObject({ anchor: 12, focus: 21 });
  const up = navigateGeometry(index, navigateGeometry(index, first, 'down', false)!, 'up', false)!;
  expect(navigateGeometry(index, up, 'up', false)?.focus).toBe(8);
});

test('wrapped_affinity', () => {
  const index = indexGeometry(list('alpha beta', [line(0, 5, 0), line(5, 10, 20)]));
  const upstream = { ...selection, anchor: 5, focus: 5, affinity: 'upstream' as const };
  expect(navigateGeometry(index, upstream, 'lineStart', false)?.focus).toBe(0);
  const end = navigateGeometry(index, { ...upstream, focus: 2, anchor: 2 }, 'lineEnd', false)!;
  expect(end.affinity).toBe('upstream');
  expect(navigateGeometry(index, end, 'lineStart', false)?.focus).toBe(0);
  expect(navigateGeometry(index, { ...end, affinity: 'downstream' }, 'lineEnd', false)?.focus).toBe(10);
  expect(hitGeometry(indexHitGeometry(list('alpha beta', [line(0, 5, 0), line(5, 10, 20)])),
    { x: 50, y: 10 })).toMatchObject({ position: 5, affinity: 'upstream' });
});

for (const [name, lines, point, position] of [
  ['empty cell', [line(0, 0, 10)], { x: 20, y: 80 }, 0],
  ['top padding', [line(0, 2, 40)], { x: 10, y: 15 }, 1],
  ['bottom padding', [line(0, 2, 10)], { x: 20, y: 80 }, 2],
  ['line gap', [line(0, 2, 10), line(3, 5, 70)], { x: 10, y: 60 }, 4],
] as const) {
  test(`table ${name} hit uses the nearest line in full and compact geometry`, () => {
    const source = list('', [...lines]);
    const cell = source.primitives[0] as TextBoxPrimitive;
    cell.h = 100;
    source.primitives = [{ kind: 'table', objectId: 2, shapeId: cell.shapeId,
      x: 0, y: 0, w: 100, h: 100, name: 'Table', label: '', primitives: [cell] }];
    for (const index of [indexGeometry, indexHitGeometry]) {
      expect(hitGeometry(index(source), point, undefined, 1)).toEqual({ kind: 'text',
        shapeId: cell.shapeId!, storyId: cell.storyId!, position,
        affinity: position === lines[lines.length - 1]!.end ? 'upstream' : 'downstream' });
      expect(hitGeometry(index(source), { x: 1, y: 50 }, undefined, 1)).toEqual({ kind: 'shape', shapeId: cell.shapeId! });
    }
  });
}

test('top-level text box padding hit keeps shape selection', () => {
  const source = list('ab', [line(0, 2, 20)]);
  for (const index of [indexGeometry, indexHitGeometry]) {
    expect(hitGeometry(index(source), { x: 20, y: 10 }, undefined, 1)).toEqual({ kind: 'shape', shapeId: 'box' });
    expect(hitGeometry(index(source), { x: 20, y: 50 }, undefined, 1)).toEqual({ kind: 'shape', shapeId: 'box' });
  }
});

test('modifier_edges_worker', () => {
  const index = indexHitGeometry(list());
  const at = { ...selection, anchor: 2, focus: 2 };
  const move = (motion: string, extend = false) => Reflect.apply(navigateGeometry, undefined,
    [index, at, motion, extend, 'alpha beta\nx\nalpha beta']) as PresentationCommandSelection;
  expect(move('wordRight', true)).toMatchObject({ anchor: 2, focus: 5 });
  expect(move('wordLeft').focus).toBe(0);
  expect(move('storyStart').focus).toBe(0);
  expect(move('storyEnd').focus).toBe(23);
});

test('pending_unicode_stops', () => {
  const json = spyOn(wasm, 'presentationDisplayListJson').mockReturnValue(JSON.stringify(list('e\u0301😀a', [line(0, 5, 0)])));
  try {
    const frame = buildCommandFrame({} as wasm.PresentationHandle, { slideId: 'slide', slideIndex: 0,
      slideRevision: 1, coversThroughSeq: 1, documentEpoch: 1, incarnation: 'worker', epoch: 1, version: 'v1' }, new Set()).value;
    const box = decodeCommandDisplayList(frame).primitives[0] as TextBoxPrimitive;
    expect(box.lines[0]!.caretStops.map((stop) => stop.position)).toEqual([0, 2, 4, 5]);
    const index = indexHitGeometry(decodeCommandDisplayList(frame));
    const at = { ...selection, anchor: 4, focus: 4 };
    expect(navigateGeometry(index, at, 'left', false)?.focus).toBe(2);
    expect(navigateGeometry(index, { ...at, anchor: 2, focus: 2 }, 'left', false)?.focus).toBe(0);
  } finally { json.mockRestore(); }
});

test('discarded_frame_keeps_media_pending_for_the_next_frame', () => {
  const displayList = list();
  displayList.primitives = [{ kind: 'image', objectId: 1, name: 'picture', assetId: 'picture', x: 0, y: 0, w: 120, h: 80 }];
  const bytes = new Uint8Array([137, 80, 78, 71]);
  const handle = { mediaBytes: () => bytes } as unknown as wasm.PresentationHandle;
  const sentMedia = new Set<string>();
  const tags = { slideId: 'slide', slideIndex: 0, slideRevision: 1, coversThroughSeq: 1,
    documentEpoch: 1, incarnation: 'worker', epoch: 1, version: 'v1' };
  const discarded = buildCommandFrame(handle, tags, sentMedia, undefined, displayList);
  expect(discarded.value.media).toHaveLength(1);
  const delivered = buildCommandFrame(handle, { ...tags, epoch: 2 }, sentMedia, undefined, displayList);
  expect(delivered.value.media).toEqual([{ assetId: 'picture', bytes: bytes.buffer }]);
  expect(sentMedia.size).toBe(0);
  for (const { assetId } of delivered.value.media) sentMedia.add(assetId);
  expect(buildCommandFrame(handle, { ...tags, epoch: 3 }, sentMedia, undefined, displayList).value.media).toEqual([]);
});

test('line_edges_respect_wrapped_graphemes', () => {
  const text = 'a😀b';
  const json = spyOn(wasm, 'presentationDisplayListJson').mockReturnValue(JSON.stringify(list(text,
    [line(0, 2, 0), line(2, 4, 20)])));
  try {
    const frame = buildCommandFrame({} as wasm.PresentationHandle, { slideId: 'slide', slideIndex: 0,
      slideRevision: 1, coversThroughSeq: 1, documentEpoch: 1, incarnation: 'worker', epoch: 1, version: 'v1' }, new Set()).value;
    for (const index of [indexGeometry, indexHitGeometry]) {
      const geometry = index(decodeCommandDisplayList(frame));
      for (const storyText of [undefined, text]) {
        expect(navigateGeometry(geometry, { ...selection, anchor: 1, focus: 1 }, 'lineEnd', false, storyText)?.focus).toBe(1);
        expect(navigateGeometry(geometry, { ...selection, anchor: 3, focus: 3 }, 'lineStart', false, storyText)?.focus).toBe(3);
      }
    }
  } finally { json.mockRestore(); }
});

test('confirmed_and_pending_ligature_stops', () => {
  const ligature = { ...line(0, 3, 0), caretStops: [
    { position: 0, x: 0 }, { position: 2, x: 20 }, { position: 3, x: 30 },
  ] };
  const json = spyOn(wasm, 'presentationDisplayListJson').mockReturnValue(JSON.stringify(list('fia', [ligature])));
  try {
    const frame = buildCommandFrame({} as wasm.PresentationHandle, { slideId: 'slide', slideIndex: 0,
      slideRevision: 1, coversThroughSeq: 1, documentEpoch: 1, incarnation: 'worker', epoch: 1, version: 'v1' }, new Set()).value;
    const index = indexHitGeometry(decodeCommandDisplayList(frame));
    for (const [focus, motion, expected] of [[0, 'right', 1], [1, 'right', 2], [2, 'left', 1], [1, 'left', 0]] as const) {
      const at = { ...selection, anchor: focus, focus };
      expect(navigateGeometry(index, at, motion, false)?.focus).toBe(expected);
      expect(navigateGeometry(index, at, motion, false, 'fia')?.focus).toBe(expected);
    }
  } finally { json.mockRestore(); }
});

test('text columns hit and selection geometry use the second column', () => {
  const second = line(5, 10, 0);
  second.columnOffset = 60;
  second.x = 60;
  second.caretStops = second.caretStops.map((stop) => ({ ...stop, x: stop.x + 60 }));
  const frame = list('alpha beta', [line(0, 5, 0), second]);
  const point = { x: 70, y: 10 };
  for (const index of [indexGeometry(frame), indexHitGeometry(frame)]) {
    expect(hitGeometry(index, point)).toMatchObject({ position: 6 });
    expect(hitGeometry(index, point, selection)).toMatchObject({ position: 6 });
  }
  expect(selectionGeometry(indexGeometry(frame), { ...selection, anchor: 6, focus: 8 }))
    .toMatchObject({ caret: { x: 90, y: 0, width: 1, height: 20 },
      selectionRects: [{ x: 70, y: 0, width: 20, height: 20 }] });
});

test('single-column drag hits keep the row beside a short line', () => {
  const frame = list('', [line(0, 1, 0), line(2, 22, 20)]);
  for (const index of [indexGeometry(frame), indexHitGeometry(frame)]) {
    expect(hitGeometry(index, { x: 150, y: 10 }, selection)).toMatchObject({ position: 1 });
  }
});


test('vertical caret movement retains its goal when starting in a later column', () => {
  const lines = [line(0, 5, 0), line(6, 11, 20), line(12, 17, 0), line(18, 23, 20)];
  for (const row of lines.slice(2)) {
    row.columnOffset = 100;
    row.x += 100;
    row.caretStops = row.caretStops.map((stop) => ({ ...stop, x: stop.x + 100 }));
  }
  const index = indexHitGeometry(list('', lines));
  const start = { ...selection, anchor: 15, focus: 15 };
  const up = navigateGeometry(index, start, 'up', false)!;
  expect(up.focus).toBe(9);
  expect(navigateGeometry(index, up, 'up', false)?.focus).toBe(3);
  expect(navigateGeometry(index, up, 'down', false)?.focus).toBe(15);
});

test('vertical caret geometry preserves centered short line column origins', () => {
  for (const offset of [100, -100]) {
    const rows = [line(0, 10, 0), line(11, 13, 20), line(14, 16, 0)];
    for (const [index, row] of rows.entries()) {
      const columnOffset = index === 2 ? offset : 0;
      row.columnOffset = columnOffset;
      row.x += columnOffset + (index === 0 ? 0 : 40);
      row.caretStops = row.caretStops.map((stop) => ({ ...stop, x: stop.x + row.x }));
    }
    for (const index of [indexGeometry(list('', rows)), indexHitGeometry(list('', rows))]) {
      const start = { ...selection, anchor: 12, focus: 12 };
      const down = navigateGeometry(index, start, 'down', false)!;
      expect(down.focus).toBe(15);
      expect(navigateGeometry(index, { ...down, goalX: undefined }, 'up', false)?.focus).toBe(12);
    }
  }
});

test('Tab motions select the next or previous table cell in full and compact geometry', () => {
  const source = list();
  const box = source.primitives[0] as TextBoxPrimitive;
  const cells = ['A1', 'B1', '', 'B2'].map((text, index): TextBoxPrimitive => ({ ...box, shapeId: 'table',
    storyId: `cell-${index}`, x: (index % 2) * 50, y: Math.floor(index / 2) * 50, w: 50, h: 50,
    lines: [line(0, text.length, Math.floor(index / 2) * 50)],
    paragraphs: [{ level: 0, runs: [{ text, fontFamily: 'Arial', fontSizePt: 12, color: '#000000' }] }] }));
  source.primitives = [box, { kind: 'table', objectId: 2, shapeId: 'table', x: 0, y: 0, w: 100, h: 100,
    name: 'Table', label: '', primitives: cells }];
  const at = (storyId: string): PresentationCommandSelection => ({ slideId: 'slide', shapeId: 'table', storyId,
    anchor: 1, focus: 1, affinity: 'downstream' });
  for (const index of [indexGeometry, indexHitGeometry]) {
    const geometry = index(source);
    expect(navigateGeometry(geometry, at('cell-0'), 'nextCell', false)).toEqual({ slideId: 'slide', shapeId: 'table',
      storyId: 'cell-1', anchor: 0, focus: 2, affinity: 'upstream' });
    expect(navigateGeometry(geometry, at('cell-1'), 'nextCell', true)).toEqual({ slideId: 'slide', shapeId: 'table',
      storyId: 'cell-2', anchor: 0, focus: 0, affinity: 'downstream' });
    expect(navigateGeometry(geometry, at('cell-2'), 'previousCell', false)).toMatchObject({ storyId: 'cell-1', anchor: 0, focus: 2 });
    expect(navigateGeometry(geometry, at('cell-3'), 'nextCell', false)).toEqual(at('cell-3'));
    expect(navigateGeometry(geometry, at('cell-0'), 'previousCell', false)).toEqual(at('cell-0'));
    expect(navigateGeometry(geometry, selection, 'nextCell', false)).toBeNull();
  }
});

for (const trailing of ['\n', '\t']) {
  test(`Tab selects cell text beyond the rendered lines with trailing ${JSON.stringify(trailing)}`, () => {
    const source = list('B1' + trailing, [line(0, 1, 0), line(1, 2, 20)]);
    const destination = source.primitives[0] as TextBoxPrimitive;
    const first = { ...destination, storyId: 'first' };
    source.primitives = [{ kind: 'table', objectId: 2, shapeId: 'box', x: 0, y: 0, w: 100, h: 100,
      name: 'Table', label: '', primitives: [first, destination] }];
    for (const index of [indexGeometry, indexHitGeometry]) {
      const geometry = index(source);
      expect(navigateGeometry(geometry, { ...selection, storyId: 'first' }, 'nextCell', false))
        .toMatchObject({ storyId: 'story', anchor: 0, focus: 3, affinity: 'upstream' });
      expect(navigateGeometry(geometry, selection, 'previousCell', false))
        .toMatchObject({ storyId: 'first', anchor: 0, focus: 3 });
    }
  });
}
