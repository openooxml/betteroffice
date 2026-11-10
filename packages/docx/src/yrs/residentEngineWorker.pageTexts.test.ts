import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import type { DocxDocumentRead, DocxDocumentReadValues } from './residentEngineWorkerProtocol';

const source = readFileSync(fileURLToPath(new URL('./residentEngineWorker.ts', import.meta.url)), 'utf8');
const start = source.indexOf("      case 'pageTexts':");
const body = source.slice(start, source.indexOf("      case 'adjacentRevision':", start));
const read = new Function('request', 'retainedFrame', 'retainedLayoutVersion', 'engine', 'layoutRevision', 'session',
  `let value; switch (request.read.kind) {${body}} return value;`) as (
    request: { read: Extract<DocxDocumentRead, { kind: 'pageTexts' }> },
    frame: { frameEpoch: number } | null, layoutVersion: string,
    engine: { version(): string }, layoutRevision: number, session: { pageTextsJson(pages: Uint32Array): string },
  ) => DocxDocumentReadValues['pageTexts'];

test('pageTexts reads retained unbuilt pages with their frame and layout stamps', () => {
  const calls: number[][] = [];
  const pages = [{ pageIndex: 15, paragraphs: [{ text: 'Heading', level: 0 }, { text: 'Table text', inTable: true as const }] }];
  const value = read({ read: { kind: 'pageTexts', pages: [15], expectFrameEpoch: 7, expectLayoutRevision: 3 } },
    { frameEpoch: 7 }, 'v1', { version: () => 'v1' }, 3,
    { pageTextsJson(input) { expect(input).toBeInstanceOf(Uint32Array); calls.push([...input]); return JSON.stringify(pages); } });
  expect(calls).toEqual([[15]]);
  expect(value).toEqual({ pages, frameEpoch: 7, layoutRevision: 3 });
});

test('pageTexts uses the neighbouring reads null staleness convention without calling the engine', () => {
  const session = { pageTextsJson() { throw new Error('Stale read reached the engine'); } };
  const request = { kind: 'pageTexts' as const, pages: [15] };
  const engine = { version: () => 'v1' };
  expect(read({ read: request }, null, 'v1', engine, 3, session)).toBeNull();
  expect(read({ read: request }, { frameEpoch: 7 }, 'old', engine, 3, session)).toBeNull();
  expect(read({ read: { ...request, expectFrameEpoch: 6 } }, { frameEpoch: 7 }, 'v1', engine, 3, session)).toBeNull();
  expect(read({ read: { ...request, expectLayoutRevision: 2 } }, { frameEpoch: 7 }, 'v1', engine, 3, session)).toBeNull();
});
