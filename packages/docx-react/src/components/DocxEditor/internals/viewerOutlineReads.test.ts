import { expect, test } from 'bun:test';
import type { ResidentDocumentRead, ResidentEngineWorkerClient, DocxOutlineHeading } from '@betteroffice/docx/yrs';
import { ViewerOutlineReads } from './viewerOutlineReads';

const headings: DocxOutlineHeading[] = [
  { story: 'body', paraId: 'h1', text: 'First heading', level: 0, position: 1 },
  { story: 'body', paraId: 'h2', text: 'Last heading', level: 1, position: 90 },
];

test('viewer headings map to outline positions and clicks at their version scroll without a read', async () => {
  const requests: ResidentDocumentRead[] = [];
  const read = (async (request: ResidentDocumentRead) => {
    requests.push(request);
    return { version: 'A', value: headings };
  }) as ResidentEngineWorkerClient['documentRead'];
  const reads = new ViewerOutlineReads(read);
  const collected = await reads.collect('A', () => 'A');
  expect(collected).toEqual([
    { text: 'First heading', level: 0, pmPos: 0 },
    { text: 'Last heading', level: 1, pmPos: 89 },
  ]);
  expect(await reads.collect('A', () => 'A')).toBe(collected!);
  const positions: number[] = [];
  await reads.navigate(89, () => 'A', (position) => { positions.push(position); });
  await reads.navigate(50, () => 'A', (position) => { positions.push(position); });
  expect(requests).toEqual([{ kind: 'headings', expectVersion: 'A' }]);
  expect(positions).toEqual([89, 50]);
});

test('a click on a heading from an older version navigates through the worker', async () => {
  const requests: ResidentDocumentRead[] = [];
  const read = (async (request: ResidentDocumentRead) => {
    requests.push(request);
    return request.kind === 'headings'
      ? { version: 'A', value: headings }
      : { version: 'B', value: { loc: { story: 'body', paraId: 'h2', offset: 0 }, position: 101 } };
  }) as ResidentEngineWorkerClient['documentRead'];
  const reads = new ViewerOutlineReads(read);
  await reads.collect('A', () => 'A');
  const positions: number[] = [];
  await reads.navigate(89, () => 'B', (position) => { positions.push(position); });
  expect(requests).toEqual([
    { kind: 'headings', expectVersion: 'A' },
    { kind: 'navigationTarget', story: 'body', paraId: 'h2' },
  ]);
  expect(positions).toEqual([100]);
});

test('a heading reply for a superseded frame is dropped', async () => {
  let complete!: (reply: { version: string; value: DocxOutlineHeading[] }) => void;
  const read = (() => new Promise<unknown>((resolve) => { complete = resolve; })) as ResidentEngineWorkerClient['documentRead'];
  const reads = new ViewerOutlineReads(read);
  let version = 'A';
  const pending = reads.collect(version, () => version);
  version = 'B';
  complete({ version: 'A', value: headings });
  expect(await pending).toBeNull();
});

test('refreshes at one version share one read and a later version supersedes it', async () => {
  const replies: Array<(reply: { version: string; value: DocxOutlineHeading[] }) => void> = [];
  const requests: ResidentDocumentRead[] = [];
  const read = ((request: ResidentDocumentRead) => {
    requests.push(request);
    return new Promise<unknown>((resolve) => { replies.push(resolve); });
  }) as ResidentEngineWorkerClient['documentRead'];
  const reads = new ViewerOutlineReads(read);
  let version = 'A';
  const first = reads.collect('A', () => version);
  const second = reads.collect('A', () => version);
  expect(requests).toHaveLength(1);
  replies[0]!({ version: 'A', value: headings });
  expect(await first).toEqual(await second);
  expect(await first).toHaveLength(2);
  version = 'B';
  const older = reads.collect('B', () => version);
  const newer = reads.collect('C', () => 'C');
  version = 'C';
  replies[1]!({ version: 'B', value: headings });
  replies[2]!({ version: 'C', value: [] });
  expect(await older).toBeNull();
  expect(await newer).toEqual([]);
});

test('a stale navigation reply cannot scroll the new frame', async () => {
  let complete!: (reply: { version: string; value: { loc: { story: string; paraId: string; offset: number }; position: number } }) => void;
  const read = ((request: ResidentDocumentRead) => request.kind === 'headings'
    ? Promise.resolve({ version: 'A', value: headings })
    : new Promise<unknown>((resolve) => { complete = resolve; })) as ResidentEngineWorkerClient['documentRead'];
  const reads = new ViewerOutlineReads(read);
  let version = 'A';
  await reads.collect(version, () => version);
  version = 'B';
  const positions: number[] = [];
  const pending = reads.navigate(89, () => version, (position) => { positions.push(position); });
  version = 'C';
  complete({ version: 'B', value: { loc: { story: 'body', paraId: 'h2', offset: 0 }, position: 101 } });
  await pending;
  expect(positions).toEqual([]);
});

test('a newer outline click supersedes a pending navigation', async () => {
  const pending: Array<(reply: { version: string; value: { loc: { story: string; paraId: string; offset: number }; position: number } }) => void> = [];
  const read = ((request: ResidentDocumentRead) => request.kind === 'headings'
    ? Promise.resolve({ version: 'A', value: headings })
    : new Promise<unknown>((resolve) => { pending.push(resolve); })) as ResidentEngineWorkerClient['documentRead'];
  const reads = new ViewerOutlineReads(read);
  await reads.collect('A', () => 'A');
  const positions: number[] = [];
  const scroll = (position: number) => { positions.push(position); };
  const first = reads.navigate(0, () => 'B', scroll);
  const second = reads.navigate(89, () => 'B', scroll);
  pending[1]!({ version: 'B', value: { loc: { story: 'body', paraId: 'h2', offset: 0 }, position: 90 } });
  await second;
  pending[0]!({ version: 'B', value: { loc: { story: 'body', paraId: 'h1', offset: 0 }, position: 1 } });
  await first;
  expect(positions).toEqual([89]);
});
