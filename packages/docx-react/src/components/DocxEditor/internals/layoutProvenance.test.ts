import { expect, test } from 'bun:test';
import { presentedWorkerFrame, presentedWorkerVersion, stampWorkerFrameVersion } from './layoutProvenance';

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
