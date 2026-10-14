import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative, resolve } from 'node:path';

import { preloadEditWasm } from '../wasm/edit';
import type { DocxContentControlQuery, DocxContentControlsOptions, DocxContentControlsResult } from './contentControls';
import { createYrsSession, type YrsSession } from './index';
import { createProposalRegistry, type DocxProposalInput } from './proposals';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';

const ROOT = resolve(import.meta.dir, '../../../..');
const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FIXTURES = [
  'crates/docx-edit/tests/fixtures',
  'crates/betteroffice-docx/tests/corpus/fixtures',
  'packages/docx/src/yrs/__fixtures__',
];
const OPTIONS: DocxContentControlsOptions[] = [
  {},
  { stories: ['body', 'headers', 'footers', 'footnotes', 'endnotes', 'comments'] },
  { maxControls: 1 },
  { maxBytes: 64 },
];
const TIMEOUT = Number(process.env.CONTENT_CONTROL_PARITY_TIMEOUT_MS ?? 60_000);
let nextClientId = 98300;
let controlsDocuments = 0;
let checkedDocuments = 0;

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));
afterAll(() => console.log(`Content-control parity: ${controlsDocuments}/${checkedDocuments} documents had controls.`));

function documents(): string[] {
  const roots = process.env.CONTENT_CONTROL_PARITY_DOCS
    ? [resolve(process.env.CONTENT_CONTROL_PARITY_DOCS)]
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

function normalize(reply: DocxContentControlsResult): unknown {
  return { ...reply, version: '<version>' };
}

function difference(worker: unknown, main: unknown, path = '$'): string | undefined {
  if (JSON.stringify(worker) === JSON.stringify(main)) return undefined;
  if (worker && main && typeof worker === 'object' && typeof main === 'object') {
    const left = worker as Record<string, unknown>;
    const right = main as Record<string, unknown>;
    const keys = Object.keys(left);
    if (JSON.stringify(keys) !== JSON.stringify(Object.keys(right))) {
      return `${path} keys: worker=${JSON.stringify(keys)}, main=${JSON.stringify(Object.keys(right))}`;
    }
    for (const key of keys) {
      const found = difference(left[key], right[key], `${path}.${key}`);
      if (found) return found;
    }
  }
  return `${path}: worker=${JSON.stringify(worker)}, main=${JSON.stringify(main)}`;
}

function compare(worker: ResidentEngineSession, main: YrsSession, label: string, failures: string[]): void {
  const full = worker.listContentControls();
  if (!full.ok) throw new Error(full.failure.message);
  const queries: DocxContentControlQuery[] = [{ kind: 'tag', tag: '<no-matching-control>' }];
  for (const control of full.content.controls) {
    if (control.tag !== null) queries.push({ kind: 'tag', tag: control.tag });
    if (control.ooxmlId !== null) queries.push({ kind: 'ooxmlId', ooxmlId: control.ooxmlId });
  }
  const check = (workerReply: DocxContentControlsResult, mainReply: DocxContentControlsResult, read: string) => {
    const found = difference(normalize(workerReply), normalize(mainReply));
    if (found) failures.push(`${label} ${read} ${found}`);
  };
  for (const options of OPTIONS) {
    check(worker.listContentControls(options), main.listContentControls(options), `list ${JSON.stringify(options)}`);
    for (const query of queries) {
      check(worker.findContentControls(query, options), main.findContentControls(query, options),
        `find ${JSON.stringify(query)} ${JSON.stringify(options)}`);
    }
  }
}

async function replicaOf(bytes: Uint8Array, worker: ResidentEngineSession): Promise<YrsSession> {
  const replica = await createYrsSession({ clientId: nextClientId++ });
  replica.openDocx(bytes, false);
  replica.loadState(worker.encodeState());
  return replica;
}

function propose(worker: ResidentEngineSession, seeded: YrsSession): void {
  const registry = createProposalRegistry(worker.proposalEngine);
  try {
    const paragraphs = worker.paragraphIdentities().paragraphs.flatMap(
      ({ session }) => session?.story === 'body' ? [session] : []
    );
    for (const paragraph of paragraphs) {
      const proposal: DocxProposalInput = {
        id: 'insertion', paragraph, op: 'insertText', at: 'end', text: ' Added text',
        suggest: { author: 'Reviewer', date: '2026-01-01T00:00:00Z' },
      };
      const result = registry.propose({ expectVersion: worker.proposalEngine.version(), proposals: [proposal] });
      if (!result.ok) continue;
      expect(result.snapshot.proposals[0]?.changed).toBe(true);
      const anchor = seeded.paragraphIdentities().paragraphs.find(
        ({ session }) => session?.story === paragraph.story && session.paraId === paragraph.paraId
      )?.session;
      if (!anchor) throw new Error('expected the seeded paragraph');
      const main = seeded.proposeChanges({
        expectVersion: seeded.version(), proposals: [{ ...proposal, paragraph: anchor }],
      });
      if (!main.ok) throw new Error(main.failure.message);
      expect(main.snapshot.proposals[0]?.changed).toBe(true);
      return;
    }
    throw new Error('expected an editable body paragraph');
  } finally {
    registry.destroy();
  }
}

async function parity(bytes: Uint8Array, name: string): Promise<void> {
  const clientId = nextClientId++;
  const worker = await createResidentEngineSession(undefined, clientId);
  const sessions: YrsSession[] = [];
  const failures: string[] = [];
  try {
    worker.openDocx(bytes);
    const full = worker.listContentControls();
    if (!full.ok) throw new Error(full.failure.message);
    checkedDocuments += 1;
    if (full.content.controls.length > 0) controlsDocuments += 1;
    const seeded = await createYrsSession({ clientId });
    sessions.push(seeded);
    seeded.openDocx(bytes, true);
    const replica = await replicaOf(bytes, worker);
    sessions.push(replica);
    compare(worker, seeded, `${name} seeded`, failures);
    compare(worker, replica, `${name} replica`, failures);
    propose(worker, seeded);
    const editedReplica = await replicaOf(bytes, worker);
    sessions.push(editedReplica);
    compare(worker, seeded, `${name} seeded after proposal`, failures);
    compare(worker, editedReplica, `${name} replica after proposal`, failures);
    expect(failures.join('\n')).toBe('');
  } finally {
    for (const session of sessions) session.destroy();
    worker.destroy();
  }
}

describe("the worker's content-control reads equal main-thread sessions'", () => {
  for (const file of documents()) {
    const name = relative(process.env.CONTENT_CONTROL_PARITY_DOCS ?? ROOT, file);
    test(name, async () => { await parity(new Uint8Array(readFileSync(file)), name); }, TIMEOUT);
  }
});
