import { expect, test } from 'bun:test';
import {
  keyResolvedReservations,
  presentedWorkerFrame,
  presentedWorkerVersion,
  reservationsKeyOf,
  reservationsVersionOfKey,
  resolvedReservationsKey,
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
