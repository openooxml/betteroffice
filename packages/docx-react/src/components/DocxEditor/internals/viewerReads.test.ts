import { expect, test } from 'bun:test';
import type { ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import { readAt } from './viewerReads';

const request = { kind: 'selectionText', story: 'body', anchor: 1, head: 6, expectVersion: 'A' } as const;

test('a null value at the requested version is an ok read', async () => {
  const read = (async () => ({ version: 'A', value: null })) as ResidentEngineWorkerClient['documentRead'];
  expect(await readAt(read, request)).toEqual({ status: 'ok', version: 'A', value: null });
});

test('a mismatched version is superseded rather than a null value', async () => {
  const read = (async () => ({ version: 'B', value: null })) as ResidentEngineWorkerClient['documentRead'];
  expect(await readAt(read, request)).toEqual({ status: 'superseded' });
});

test('a rejected read is superseded', async () => {
  const read = (async () => { throw new Error('read failed'); }) as ResidentEngineWorkerClient['documentRead'];
  expect(await readAt(read, request)).toEqual({ status: 'superseded' });
});
