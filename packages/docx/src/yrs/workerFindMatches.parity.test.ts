import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { basename, join, resolve } from 'node:path';
import { preloadEditWasm } from '../wasm/edit';
import { findBodyMatches } from './findMatches';
import { createYrsSession, type YrsLoc, type YrsSession } from './index';
import { createYrsInputPositionMap } from './inputPositionMap';
import { createProposalRegistry } from './proposals';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';
import { residentBodyPositions } from './residentSearch';
import { createYrsPositionProjection, yrsLocToProjectedDisplayPosition } from './yrsPositionProjection';

const ROOT = resolve(import.meta.dir, '../../../..');
const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FIXTURES = [
  'crates/docx-edit/tests/fixtures',
  'crates/betteroffice-docx/tests/corpus/fixtures',
  'packages/docx/src/yrs/__fixtures__',
];
const TIMEOUT = Number(process.env.FIND_PARITY_TIMEOUT_MS ?? 60_000);
let nextClientId = 97400;
let docs = 0;
let queries = 0;
let totalMatches = 0;

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));
afterAll(() => console.log(`Find parity: docs=${docs}, queries=${queries}, total matches=${totalMatches}`));

function documents(): string[] {
  const roots = process.env.FIND_PARITY_DOCS
    ? [resolve(process.env.FIND_PARITY_DOCS)]
    : FIXTURES.map((dir) => join(ROOT, dir));
  const found: string[] = [];
  const visit = (path: string) => {
    if (statSync(path).isDirectory()) {
      for (const entry of readdirSync(path).sort()) visit(join(path, entry));
    } else if (path.endsWith('.docx')) {
      found.push(path);
    }
  };
  for (const root of roots) visit(root);
  return found.sort();
}

function searchTexts(reader: Pick<YrsSession, 'paragraphs'>): string[] {
  const texts = reader.paragraphs('body').map(({ text }) => text).filter((text) => text.trim());
  const searches = new Set<string>();
  for (const text of texts.slice(0, 5)) {
    const word = text.match(/[\p{L}\p{N}_]+/u)?.[0] ?? text.trim().split(/\s+/)[0]!;
    searches.add(word);
    searches.add(word.toUpperCase());
  }
  const substring = texts.find((text) => text.length >= 2)?.slice(0, 2);
  if (substring) searches.add(substring);
  const phrase = texts.map((text) => text.match(/\S+ +\S+/)?.[0]).find(Boolean) ??
    texts.find((text) => text.includes(' '));
  if (phrase) searches.add(phrase);
  let missing = 'absent_find_parity_query';
  while (texts.some((text) => text.toLowerCase().includes(missing))) missing += '_';
  searches.add(missing);
  return [...searches];
}

async function replicaOf(bytes: Uint8Array, worker: ResidentEngineSession): Promise<YrsSession> {
  const replica = await createYrsSession({ clientId: nextClientId++ });
  replica.openDocx(bytes, false);
  replica.loadState(worker.encodeState());
  return replica;
}

function compare(worker: ResidentEngineSession, replica: YrsSession): void {
  const workerPositions = residentBodyPositions(worker.geometryReader);
  const projections = new Map<string, ReturnType<typeof createYrsPositionProjection>>();
  const maps = new Map<string, ReturnType<typeof createYrsInputPositionMap>>();
  const replicaPositions = (loc: YrsLoc) => yrsLocToProjectedDisplayPosition(
    replica,
    (root) => {
      if (!projections.has(root)) projections.set(root, createYrsPositionProjection(replica, root));
      return projections.get(root)!;
    },
    loc,
    'body',
    (story) => {
      if (!maps.has(story)) maps.set(story, createYrsInputPositionMap(story, replica.paragraphSpans(story)));
      return maps.get(story)!;
    }
  );
  for (const searchText of searchTexts(worker.geometryReader)) {
    for (const matchCase of [false, true]) {
      for (const matchWholeWord of [false, true]) {
        const options = { matchCase, matchWholeWord };
        const matches = findBodyMatches(worker.geometryReader, workerPositions, searchText, options);
        const expected = findBodyMatches(replica, replicaPositions, searchText, options);
        expect(matches).toEqual(expected);
        queries += 1;
        totalMatches += matches.length;
      }
    }
  }
}

function propose(worker: ResidentEngineSession): void {
  const paragraph = worker.paragraphIdentities().paragraphs
    .find(({ session }) => session?.story === 'body')?.session;
  if (!paragraph) throw new Error('expected a body paragraph');
  const registry = createProposalRegistry(worker.proposalEngine);
  try {
    const result = registry.propose({
      expectVersion: worker.proposalEngine.version(),
      proposals: [{
        id: 'insertion',
        paragraph,
        suggest: { author: 'Reviewer', date: '2026-01-01T00:00:00Z' },
        op: 'insertText', at: 'end', text: ' Added text',
      }],
    });
    if (!result.ok) throw new Error(result.failure.message);
    expect(result.snapshot.proposals).toHaveLength(1);
    expect(result.snapshot.proposals[0]).toMatchObject({ changed: true });
    expect(result.snapshot.proposals[0]!.revisionIds.length).toBeGreaterThan(0);
  } finally {
    registry.destroy();
  }
}

async function parity(bytes: Uint8Array, afterProposal = false): Promise<void> {
  const worker = await createResidentEngineSession(undefined, nextClientId++);
  let replica: YrsSession | undefined;
  try {
    worker.openDocx(bytes);
    replica = await replicaOf(bytes, worker);
    compare(worker, replica);
    if (afterProposal) {
      propose(worker);
      replica.destroy();
      replica = await replicaOf(bytes, worker);
      compare(worker, replica);
    } else {
      docs += 1;
    }
  } finally {
    replica?.destroy();
    worker.destroy();
  }
}

describe('worker find matches equal the viewer replica', () => {
  const files = documents();
  files.forEach((file, index) => {
    test(`document ${index + 1}`, async () => {
      await parity(new Uint8Array(readFileSync(file)));
    }, TIMEOUT);
  });
  test('after a worker proposal', async () => {
    const file = files.find((file) => basename(file) === 'pages.docx') ?? files[0];
    expect(file).toBeDefined();
    await parity(new Uint8Array(readFileSync(file!)), true);
  }, TIMEOUT);
});
