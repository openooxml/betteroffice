import { expect, test } from 'bun:test';
import { HEADER_BUDGET_BYTES, headerBytes, isRenderFrameHeader, surfaceKeyId, transferList,
  type FrameHeader, type RenderFrame, type SurfaceKey } from './types';

const header: FrameHeader = { session: 'session', incarnation: 1, epoch: 8, baseEpoch: 7,
  coversThroughSeq: 42, viewportGeneration: 3, scale: 1.25, retire: [], progressive: false,
  geometryRevision: { 'page:1': 2 } };

test('header validation accepts declared fields and rejects nested graphs and wrong types', () => {
  expect(isRenderFrameHeader(header)).toBe(true);
  expect(isRenderFrameHeader({ ...header, retire: [{ kind: 'page', id: 'old' }] })).toBe(true);
  for (const invalid of [null, [], { ...header, epoch: 1.2 }, { ...header, incarnation: -1 },
    { ...header, coversThroughSeq: '42' }, { ...header, baseEpoch: NaN }, { ...header, scale: Infinity },
    { ...header, scale: 0 }, { ...header, progressive: 1 }, { ...header, viewportGeneration: {} },
    { ...header, session: {} }, { ...header, retire: [{}] },
    { ...header, retire: [{ kind: 'other', id: '1' }] },
    { ...header, retire: [{ kind: 'page', id: '1', content: {} }] },
    { ...header, geometryRevision: { page: {} } }, { ...header, extra: { content: [] } }]) {
    expect(isRenderFrameHeader(invalid)).toBe(false);
  }
});

test('ten-surface header fits the budget and measures UTF-8 without frame content', () => {
  const keys: SurfaceKey[] = Array.from({ length: 10 }, (_, index) => ({ kind: 'pane', id: `surface:${index}` }));
  const realistic = { ...header, retire: keys, geometryRevision: Object.fromEntries(keys.map((key) => [surfaceKeyId(key), 20])) };
  expect(headerBytes(realistic)).toBeLessThanOrEqual(HEADER_BUDGET_BYTES);
  expect(headerBytes({ ...header, session: 'é' })).toBe(headerBytes({ ...header, session: 'a' }) + 1);
  const withContent = { ...realistic, geometry: new ArrayBuffer(4096), patches: [] };
  expect(headerBytes(withContent)).toBe(headerBytes(realistic));
  expect(surfaceKeyId({ kind: 'page', id: '1' })).not.toBe(surfaceKeyId({ kind: 'slide', id: '1' }));
});

test('transfer list contains geometry, echo, patch and plate bitmaps in order', () => {
  const bitmap = () => ({ width: 1, height: 1, close() {} });
  const a = bitmap(), b = bitmap(), c = bitmap();
  const surface: SurfaceKey = { kind: 'page', id: '1' }, rect = { x: 0, y: 0, width: 1, height: 1 };
  const frame: RenderFrame<typeof a> = { ...header, geometry: new ArrayBuffer(1), echo: new ArrayBuffer(2),
    patches: [{ surface, rect, bitmap: a, full: true }, { surface, rect, bitmap: b, full: false }],
    plates: [{ surface, rect, bitmap: c }] };
  expect(transferList(frame)).toEqual([frame.geometry!, frame.echo!, a, b, c]);
  expect(transferList({ ...header, patches: [] })).toEqual([]);
  expect(transferList({ ...header, patches: frame.patches })).toEqual([a, b]);
});
