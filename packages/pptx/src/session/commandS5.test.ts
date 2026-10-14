import { expect, spyOn, test } from 'bun:test';
import type { PositionedTextLine, SlideDisplayList, TextBoxPrimitive } from '../types';
import * as wasm from '../wasm/loader';
import { buildCommandFrame } from './commandFrame';
import { decodeCommandDisplayList, hitGeometry, indexGeometry, indexHitGeometry, navigateGeometry } from './commandGeometry';
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
