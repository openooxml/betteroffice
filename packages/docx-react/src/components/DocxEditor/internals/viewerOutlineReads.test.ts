import { expect, test } from 'bun:test';
import type { ResidentDocumentRead, ResidentEngineWorkerClient, DocxOutlineHeading } from '@betteroffice/docx/yrs';
import { ViewerOutlineReads } from './viewerOutlineReads';

const headings: DocxOutlineHeading[] = [
  { story: 'body', paraId: 'h1', text: 'First heading', level: 0, position: 1 },
  { story: 'body', paraId: 'h2', text: 'Last heading', level: 1, position: 90 },
];

test('viewer headings map to outline positions and clicks navigate through the worker', async () => {
  const requests: ResidentDocumentRead[] = [];
  const read = (async (request: ResidentDocumentRead) => {
    requests.push(request);
    return { version: 'A', value: request.kind === 'headings' ? headings : {
      loc: { story: 'body', paraId: 'h2', offset: 0 }, position: 101,
    } };
  }) as ResidentEngineWorkerClient['documentRead'];
  const reads = new ViewerOutlineReads(read);
  expect(await reads.collect('A', () => 'A')).toEqual([
    { text: 'First heading', level: 0, pmPos: 0 },
    { text: 'Last heading', level: 1, pmPos: 89 },
  ]);
  const positions: number[] = [];
  await reads.navigate(89, () => 'A', (position) => { positions.push(position); });
  await reads.navigate(50, () => 'A', (position) => { positions.push(position); });
  expect(requests).toEqual([
    { kind: 'headings', expectVersion: 'A' },
    { kind: 'navigationTarget', story: 'body', paraId: 'h2' },
  ]);
  expect(positions).toEqual([100, 50]);
});

test('a heading reply for a superseded frame is dropped', async () => {
  let complete!: (reply: { version: string; value: DocxOutlineHeading[] }) => void;
  const read = (() => new Promise((resolve) => { complete = resolve; })) as ResidentEngineWorkerClient['documentRead'];
  const reads = new ViewerOutlineReads(read);
  let version = 'A';
  const pending = reads.collect(version, () => version);
  version = 'B';
  complete({ version: 'A', value: headings });
  expect(await pending).toBeNull();
});

test('only the latest refresh is applied even at the same version', async () => {
  const replies: Array<(reply: { version: string; value: DocxOutlineHeading[] }) => void> = [];
  const read = (() => new Promise((resolve) => { replies.push(resolve); })) as ResidentEngineWorkerClient['documentRead'];
  const reads = new ViewerOutlineReads(read);
  const first = reads.collect('A', () => 'A');
  const second = reads.collect('A', () => 'A');
  replies[1]!({ version: 'A', value: headings });
  expect(await second).not.toBeNull();
  replies[0]!({ version: 'A', value: [] });
  expect(await first).toBeNull();
});

test('a stale navigation reply cannot scroll the new frame', async () => {
  let complete!: (reply: { version: string; value: { loc: { story: string; paraId: string; offset: number }; position: number } }) => void;
  const read = ((request: ResidentDocumentRead) => request.kind === 'headings'
    ? Promise.resolve({ version: 'A', value: headings })
    : new Promise((resolve) => { complete = resolve; })) as ResidentEngineWorkerClient['documentRead'];
  const reads = new ViewerOutlineReads(read);
  let version = 'A';
  await reads.collect(version, () => version);
  const positions: number[] = [];
  const pending = reads.navigate(89, () => version, (position) => { positions.push(position); });
  version = 'B';
  complete({ version: 'A', value: { loc: { story: 'body', paraId: 'h2', offset: 0 }, position: 101 } });
  await pending;
  expect(positions).toEqual([]);
});


test('a newer outline click supersedes a pending navigation at the same version', async () => {
  const pending: Array<(reply: { version: string; value: { loc: { story: string; paraId: string; offset: number }; position: number } }) => void> = [];
  const read = ((request: ResidentDocumentRead) => request.kind === 'headings'
    ? Promise.resolve({ version: 'A', value: headings })
    : new Promise((resolve) => { pending.push(resolve); })) as ResidentEngineWorkerClient['documentRead'];
  const reads = new ViewerOutlineReads(read);
  await reads.collect('A', () => 'A');
  const positions: number[] = [];
  const scroll = (position: number) => { positions.push(position); };
  const first = reads.navigate(0, () => 'A', scroll);
  const second = reads.navigate(89, () => 'A', scroll);
  pending[1]!({ version: 'A', value: { loc: { story: 'body', paraId: 'h2', offset: 0 }, position: 90 } });
  await second;
  pending[0]!({ version: 'A', value: { loc: { story: 'body', paraId: 'h1', offset: 0 }, position: 1 } });
  await first;
  expect(positions).toEqual([89]);
});
