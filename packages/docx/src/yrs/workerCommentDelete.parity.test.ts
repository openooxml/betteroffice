import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, decodeDocxHostJson, type YrsSession } from './index';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';
import type { DocxSidebarReader } from './sidebarReads';
import { sidebarDocx } from './__fixtures__/sidebarDocx';

const ROOT = resolve(import.meta.dir, '../../../..');
const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FIXTURES = [
  'crates/docx-edit/tests/fixtures',
  'crates/betteroffice-docx/tests/corpus/fixtures',
  'packages/docx/src/yrs/__fixtures__',
];
const TIMEOUT = Number(process.env.COMMENT_DELETE_PARITY_TIMEOUT_MS ?? 60_000);
let clientId = 800_000;
let docs = 0;
let docsWithComments = 0;
let deletions = 0;

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));
afterAll(() => {
  console.log(`Comment delete parity: ${docs} docs, ${docsWithComments} docs with comments, ${deletions} deletions compared`);
});

function documents(): string[] {
  const roots = process.env.COMMENT_DELETE_PARITY_DOCS
    ? [resolve(process.env.COMMENT_DELETE_PARITY_DOCS)]
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

function anchors(reader: Pick<YrsSession, 'resolveComment'>, id: string) {
  try {
    return reader.resolveComment(id);
  } catch {
    return [];
  }
}

function equalReads(actual: DocxSidebarReader, expected: YrsSession, ids: string[]) {
  expect(actual.storyIds()).toEqual(expected.storyIds());
  for (const story of new Set([...actual.storyIds(), ...expected.storyIds()])) {
    expect(actual.paragraphs(story)).toEqual(expected.paragraphs(story));
    expect(actual.storySegments(story)).toEqual(expected.storySegments(story));
  }
  for (const id of ids) expect(anchors(actual, id)).toEqual(anchors(expected, id));
}

async function replicaOf(worker: ResidentEngineSession, bytes: Uint8Array) {
  const replica = await createYrsSession({ clientId: clientId++ });
  try {
    replica.openDocx(bytes, false);
    replica.loadState(worker.encodeState());
    return replica;
  } catch (error) {
    replica.destroy();
    throw error;
  }
}

async function parity(bytes: Uint8Array) {
  const initial = await createResidentEngineSession(undefined, clientId++);
  let ids: string[];
  try {
    const host = decodeDocxHostJson(initial.openDocx(bytes), bytes);
    ids = (host.document.package.document.comments ?? []).map(({ id }) => String(id));
    docs += 1;
  } finally {
    initial.destroy();
  }
  if (ids.length === 0) return;
  docsWithComments += 1;
  for (const id of ids.slice(0, 20)) {
    const worker = await createResidentEngineSession(undefined, clientId++);
    let replica: YrsSession | undefined;
    let handedOver: YrsSession | undefined;
    try {
      worker.openDocx(bytes);
      replica = await replicaOf(worker, bytes);
      equalReads(worker.geometryReader, replica, ids);
      try { replica.applyRawOps('body', [{ op: 'removeComment', id }]); } catch {}
      try { worker.applyRawOps('body', [{ op: 'removeComment', id }]); } catch {}
      equalReads(worker.geometryReader, replica, ids);
      expect(anchors(worker, id)).toEqual([]);
      handedOver = await replicaOf(worker, bytes);
      equalReads(handedOver, replica, ids);
      deletions += 1;
    } finally {
      handedOver?.destroy();
      replica?.destroy();
      worker.destroy();
    }
  }
}

describe('worker comment deletion equals the main-thread path', () => {
  for (const [index, file] of documents().entries()) {
    test(`document ${index + 1}`, async () => {
      await parity(new Uint8Array(readFileSync(file)));
    }, TIMEOUT);
  }
  test('synthetic comment and revisions', async () => {
    const before = deletions;
    await parity(sidebarDocx());
    expect(deletions).toBeGreaterThan(before);
  }, TIMEOUT);
});
