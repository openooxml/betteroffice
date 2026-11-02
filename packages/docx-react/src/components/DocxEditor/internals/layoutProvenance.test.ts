import { expect, test } from 'bun:test';
import {
  keyResolvedReservations,
  presentedWorkerFrame,
  presentedWorkerVersion,
  reservationsKeyOf,
  reservationsOf,
  reservationsVersionOfKey,
  resolvedReservationsKey,
  sameReservationOwner,
  stampReservationsKey,
  stampWorkerFrameVersion,
} from './layoutProvenance';

test('presented frame provenance includes preview and as-opened status from its display list', () => {
  const displayList = {};
  stampWorkerFrameVersion(displayList, 'preview', true);
  const queries = { displayList };
  expect(presentedWorkerFrame(queries)).toEqual({ version: 'preview', preview: true, asOpened: false });
  expect(presentedWorkerVersion(queries)).toBe('preview');
  stampWorkerFrameVersion(queries, 'full', false, true);
  expect(presentedWorkerFrame(queries)).toEqual({ version: 'full', preview: false, asOpened: true });
  expect(presentedWorkerVersion(queries)).toBe('full');
  stampWorkerFrameVersion(queries, 'changed');
  expect(presentedWorkerFrame(queries)).toEqual({ version: 'changed', preview: false, asOpened: false });
  expect(presentedWorkerVersion(queries)).toBe('changed');
});

test('unstamped queries have no worker frame provenance', () => {
  expect(presentedWorkerFrame({ displayList: {} })).toBeNull();
  expect(presentedWorkerFrame(null)).toBeNull();
});

test('a changed reservation set at the same version needs a new painted layout', () => {
  const first = {};
  const next = {};
  const layout = {};
  const firstKey = keyResolvedReservations(first, 'v1', 1);
  const nextKey = keyResolvedReservations(next, 'v1', 2);
  stampReservationsKey(layout, firstKey);
  expect(reservationsKeyOf(layout)).toBe(firstKey);
  expect(reservationsKeyOf(layout)).not.toBe(resolvedReservationsKey(next));
  expect(reservationsVersionOfKey(firstKey)).toBe('v1');
  expect(reservationsVersionOfKey(nextKey)).toBe('v1');
  expect(resolvedReservationsKey(null)).toBe('');
  expect(reservationsVersionOfKey('')).toBeNull();
  expect(reservationsKeyOf({})).toBeNull();
});

test('an owner whose own reservations are unchanged matches a set laid out under another key', () => {
  const entry = (owner: string, width: number) => ({ owner, id: owner, story: 'body', paraId: 'p', start: 0, end: 0, kind: 'inline' as const, width, height: 20 });
  const set = (version: string, items: { owner: string; width: number }[]) => ({
    version, sets: items.map(({ owner, width }) => ({ owner, version: 'v1', items: [{ id: owner, kind: 'inline' as const, width, height: 20, anchor: { kind: 'range' as const, version: 'v1', range: { story: 'body', view: 'accepted' as const, start: { paraId: 'p', offset: 0 }, end: { paraId: 'p', offset: 0 } } } }] })),
  });
  const laidPayload = set('v1', [{ owner: 'a', width: 35 }, { owner: 'b', width: 20 }]);
  const laid = { '["a","a"]': entry('a', 35), '["b","b"]': entry('b', 20) };
  const nextPayload = set('v2', [{ owner: 'a', width: 35 }]);
  const next = { '["a","a"]': entry('a', 35) };
  const layout = {};
  stampReservationsKey(layout, keyResolvedReservations(laid, 'v1', 1, laidPayload), laid);
  keyResolvedReservations(next, 'v2', 2, nextPayload);
  const current = { resolved: next, payload: null };
  expect(sameReservationOwner(reservationsOf(layout), current, 'a')).toBe(true);
  expect(sameReservationOwner(reservationsOf(layout), current, 'b')).toBe(false);
  const wider = { '["a","a"]': entry('a', 40) };
  keyResolvedReservations(wider, 'v2', 3, set('v2', [{ owner: 'a', width: 40 }]));
  expect(sameReservationOwner(reservationsOf(layout), { resolved: wider, payload: null }, 'a')).toBe(false);
  const viewer = set('v1', [{ owner: 'a', width: 35 }]);
  const frame = {};
  stampReservationsKey(frame, keyResolvedReservations(viewer, 'v1', 4, viewer), viewer);
  expect(sameReservationOwner(reservationsOf(frame), { payload: set('v1', [{ owner: 'a', width: 35 }]) }, 'a')).toBe(true);
  expect(sameReservationOwner(reservationsOf(frame), { payload: set('v2', [{ owner: 'a', width: 35 }]) }, 'a')).toBe(false);
  stampReservationsKey(frame, 'unknown');
  expect(reservationsOf(frame)).toBeNull();
});
