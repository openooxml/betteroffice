import { expect, test } from 'bun:test';
import { presentedWorkerFrame, presentedWorkerVersion, stampWorkerFrameVersion } from './layoutProvenance';

test('presented frame provenance includes preview status from its display list', () => {
  const displayList = {};
  stampWorkerFrameVersion(displayList, 'preview', true);
  const queries = { displayList };
  expect(presentedWorkerFrame(queries)).toEqual({ version: 'preview', preview: true });
  expect(presentedWorkerVersion(queries)).toBe('preview');
  stampWorkerFrameVersion(queries, 'full');
  expect(presentedWorkerFrame(queries)).toEqual({ version: 'full', preview: false });
  expect(presentedWorkerVersion(queries)).toBe('full');
});

test('unstamped queries have no worker frame provenance', () => {
  expect(presentedWorkerFrame({ displayList: {} })).toBeNull();
  expect(presentedWorkerFrame(null)).toBeNull();
});
