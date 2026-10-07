import { expect, test } from 'bun:test';
import {
  clearPresented,
  isPresented,
  markPresented,
  presentedWorkerFrame,
  presentedWorkerVersion,
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

test('worker canvas presentation matches the zoom even when the display list is reused', () => {
  const host = {};
  const displayList = {};
  markPresented(host, displayList, { worker: true, zoom: 1 });
  expect(isPresented(host, displayList)).toBe(true);
  expect(isPresented(host, displayList, 1)).toBe(true);
  expect(isPresented(host, displayList, 1.5)).toBe(false);

  clearPresented(host, { worker: true, zoom: 1.5 });
  expect(isPresented(host, displayList, 1.5)).toBe(false);
  markPresented(host, displayList, { worker: true, zoom: 1.5 });
  expect(isPresented(host, displayList, 1)).toBe(false);
  expect(isPresented(host, displayList, 1.5)).toBe(true);

  markPresented(host, displayList);
  expect(isPresented(host, displayList, 1)).toBe(true);
  expect(isPresented(host, displayList, 1.5)).toBe(true);
  clearPresented(host);
  expect(isPresented(host, displayList)).toBe(false);
});
