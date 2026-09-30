import { afterAll, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import {
  applyFrameDeltaOwned,
  decodeFrameDelta,
  FrameDeltaError,
  type RetainedFrame,
} from '../layout/render/frameDelta';
import { preloadEditWasm } from '../wasm/edit';
import { residentWorkerFactory, type InProcessResidentWorker } from './__fixtures__/residentWorker';
import { createYrsSession, type DocxEditRequest, type YrsSession } from './index';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(
  import.meta.dir,
  '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);
const LAYOUT = JSON.stringify({
  bodyStory: 'body',
  regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
  renderEnv: {},
});

let startWorker: () => InProcessResidentWorker;
const sessions: YrsSession[] = [];
const clients: ResidentEngineWorkerClient[] = [];

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  startWorker = await residentWorkerFactory();
});

afterAll(() => {
  for (const client of clients.splice(0)) client.destroy();
  for (const session of sessions.splice(0)) session.destroy();
});

function accepted(session: YrsSession): string[] {
  const read = session.readParagraphs({ view: 'accepted' });
  if (!read.ok) throw new Error(read.failure.message);
  return read.paragraphs.map((paragraph) => paragraph.text);
}

function frameText(frame: RetainedFrame): string {
  return frame.displayList.pages
    .flatMap((page) => page.primitives)
    .map((primitive) => (primitive.kind === 'glyphRun' || primitive.kind === 'text' ? primitive.text : ''))
    .join('');
}

test('host batches drain worker input, invalidate the worker once and never adopt stale frames', async () => {
  const main = await createYrsSession({ clientId: 5101 });
  sessions.push(main);
  const { paraId } = main.createStory('body', 'Seed');
  main.registerFont(new Uint8Array(readFileSync(FONT)));
  main.layoutDocumentWithRegionsJson(LAYOUT);
  main.setSelection({ story: 'body', paraId, offset: 4 });
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  const forwarded: Uint8Array[] = [];
  let adopting = false;
  main.onUpdate((update) => {
    if (adopting) return;
    forwarded.push(update);
    client.invalidate(update, null);
  });
  const booted = await client.bootstrap(main.residentWorkerSnapshot()!, '{}');
  let frame = applyFrameDeltaOwned(null, decodeFrameDelta(booted.frame));
  expect(frameText(frame)).toBe('Seed');

  const request = (expectVersion: string): DocxEditRequest => ({
    expectVersion,
    steps: [
      {
        op: 'replaceText',
        target: { kind: 'search', text: 'Seed', within: { kind: 'paragraph', story: 'body', paraId }, view: 'accepted' },
        text: 'Batch',
      },
    ],
  });
  const readBeforeTyping = main.version();
  const typed = await client.applyInput(' typed', main.selection()!, frame.frameEpoch);
  if (!typed.applied) throw new Error('the worker refused resident input');
  adopting = true;
  for (const update of typed.updates) main.applyLocalUpdate(update);
  adopting = false;
  const typedFrame = typed.frame.slice();
  frame = applyFrameDeltaOwned(frame, decodeFrameDelta(typed.frame));
  expect(frameText(frame)).toBe('Seed typed');

  expect(main.applyEdits(request(readBeforeTyping))).toMatchObject({
    ok: false,
    failure: { code: 'stale-version' },
  });
  expect(accepted(main)).toEqual(['Seed typed']);
  expect(forwarded).toHaveLength(0);
  expect(client.isReady()).toBe(true);

  const applied = main.applyEdits(request(main.version()));
  expect(applied).toMatchObject({ ok: true, applied: true, changedStories: ['body'] });
  expect(forwarded).toHaveLength(1);
  expect(client.isReady()).toBe(false);
  expect(await client.applyInput('!', main.selection()!, frame.frameEpoch)).toEqual({ applied: false });

  main.layoutDocumentWithRegionsJson(LAYOUT);
  const synced = await client.sync(
    main.residentWorkerSnapshot({
      knownStateVector: client.remoteStateVector(),
      knownFontsRevision: client.syncedFontsRevision(),
    })!,
    '{}',
    frame.frameEpoch
  );
  expect(client.remoteStateVector()).toEqual(main.encodeStateVector());
  frame = applyFrameDeltaOwned(frame, decodeFrameDelta(synced.frame));
  expect(frameText(frame)).toBe('Batch typed');
  let stale: unknown = null;
  try {
    applyFrameDeltaOwned(frame, decodeFrameDelta(typedFrame));
  } catch (error) {
    stale = error;
  }
  expect(stale).toBeInstanceOf(FrameDeltaError);
  expect((stale as FrameDeltaError).code).toBe('stale-frame');
});

test('one resident delete request removes several characters and stops at the story start', async () => {
  const main = await createYrsSession({ clientId: 5102 });
  sessions.push(main);
  const { paraId } = main.createStory('body', 'Alpha');
  const { secondParaId } = main.splitParagraph({ story: 'body', paraId, offset: 2 });
  main.registerFont(new Uint8Array(readFileSync(FONT)));
  main.layoutDocumentWithRegionsJson(LAYOUT);
  main.setSelection({ story: 'body', paraId: secondParaId, offset: 1 });
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  const booted = await client.bootstrap(main.residentWorkerSnapshot()!, '{}');
  let frame = applyFrameDeltaOwned(null, decodeFrameDelta(booted.frame));
  expect(frameText(frame)).toBe('Alpha');

  const remove = async (count: number) => {
    const deleted = await client.applyDelete(
      'backward',
      main.selection()!,
      frame.frameEpoch,
      false,
      false,
      count
    );
    if (!deleted.applied) throw new Error('the worker refused resident deletion');
    for (const update of deleted.updates) main.applyLocalUpdate(update);
    main.setSelection(deleted.selection!.anchor, deleted.selection!.head);
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(deleted.frame));
    return deleted.deletedUnits;
  };
  expect(await remove(3)).toBe(3);
  expect(accepted(main)).toEqual(['Aha']);
  expect(frameText(frame)).toBe('Aha');
  expect(await remove(10)).toBe(1);
  expect(accepted(main)).toEqual(['ha']);
  expect(frameText(frame)).toBe('ha');
});

test('one resident delete request merges at most one paragraph', async () => {
  const main = await createYrsSession({ clientId: 5103 });
  sessions.push(main);
  const { paraId } = main.createStory('body', 'Alpha');
  const { secondParaId } = main.splitParagraph({ story: 'body', paraId, offset: 2 });
  const { secondParaId: lastParaId } = main.splitParagraph({
    story: 'body',
    paraId: secondParaId,
    offset: 0,
  });
  main.registerFont(new Uint8Array(readFileSync(FONT)));
  main.layoutDocumentWithRegionsJson(LAYOUT);
  main.setSelection({ story: 'body', paraId: lastParaId, offset: 0 });
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  const booted = await client.bootstrap(main.residentWorkerSnapshot()!, '{}');
  let frame = applyFrameDeltaOwned(null, decodeFrameDelta(booted.frame));
  expect(accepted(main)).toEqual(['Al', '', 'pha']);

  const remove = async (count: number) => {
    const deleted = await client.applyDelete(
      'backward',
      main.selection()!,
      frame.frameEpoch,
      false,
      false,
      count
    );
    if (!deleted.applied) throw new Error('the worker refused resident deletion');
    for (const update of deleted.updates) main.applyLocalUpdate(update);
    main.setSelection(deleted.selection!.anchor, deleted.selection!.head);
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(deleted.frame));
    return deleted.deletedUnits;
  };
  expect(await remove(3)).toBe(1);
  expect(accepted(main)).toEqual(['Al', 'pha']);
  expect(await remove(3)).toBe(3);
  expect(accepted(main)).toEqual(['pha']);
  expect(frameText(frame)).toBe('pha');
});

test('typing in a table cell goes through the resident worker', async () => {
  const main = await createYrsSession({ clientId: 5104 });
  sessions.push(main);
  const { paraId } = main.createStory('body', 'Anchor');
  main.insertTable({ story: 'body', paraId, offset: 0 }, 1, 2);
  const cell = main.storyIds().find((story) => story.startsWith('body:'))!;
  const cellParagraph = main.paragraphs(cell)[0]!.paraId;
  main.insertText({ story: cell, paraId: cellParagraph, offset: 0 }, 'Cell');
  main.registerFont(new Uint8Array(readFileSync(FONT)));
  main.layoutDocumentWithRegionsJson(LAYOUT);
  main.setSelection({ story: cell, paraId: cellParagraph, offset: 4 });
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  const booted = await client.bootstrap(main.residentWorkerSnapshot()!, '{}');
  let frame = applyFrameDeltaOwned(null, decodeFrameDelta(booted.frame));
  expect(frameText(frame)).toContain('Cell');

  const typed = await client.applyInput(' typed', main.selection()!, frame.frameEpoch);
  if (!typed.applied) throw new Error('the worker refused resident input in a cell');
  for (const update of typed.updates) main.applyLocalUpdate(update);
  frame = applyFrameDeltaOwned(frame, decodeFrameDelta(typed.frame));
  expect(frameText(frame)).toContain('Cell typed');
  expect(main.paragraphs(cell)[0]!.text).toBe('Cell typed');
  expect(typed.selection?.head).toMatchObject({ story: cell, offset: 10 });
});

test('a resident delete does not merge a paragraph forward over a table', async () => {
  const main = await createYrsSession({ clientId: 5105 });
  sessions.push(main);
  const { paraId } = main.createStory('body', 'Alpha');
  const { secondParaId } = main.splitParagraph({ story: 'body', paraId, offset: 5 });
  main.insertText({ story: 'body', paraId: secondParaId, offset: 0 }, 'Omega');
  main.insertTable({ story: 'body', paraId: secondParaId, offset: 0 }, 1, 1);
  main.registerFont(new Uint8Array(readFileSync(FONT)));
  main.layoutDocumentWithRegionsJson(LAYOUT);
  main.setSelection({ story: 'body', paraId, offset: 5 });
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  const booted = await client.bootstrap(main.residentWorkerSnapshot()!, '{}');
  const frame = applyFrameDeltaOwned(null, decodeFrameDelta(booted.frame));
  const before = accepted(main);

  expect(
    await client.applyDelete('forward', main.selection()!, frame.frameEpoch, false, false)
  ).toEqual({ applied: false });
  expect(accepted(main)).toEqual(before);
  expect(client.isReady()).toBe(true);
});

test('the worker lays a host batch out exactly as the main thread does', async () => {
  const main = await createYrsSession({ clientId: 5110 });
  sessions.push(main);
  const filler = 'lorem ipsum dolor sit amet '.repeat(8);
  const ids = main.loadStories([
    {
      storyId: 'body',
      paragraphs: Array.from({ length: 240 }, (_, index) => ({ text: `${index} ${filler}` })),
    },
  ]).body!;
  main.registerFont(new Uint8Array(readFileSync(FONT)));
  main.adoptResidentWorkerLayout!(LAYOUT);
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  const layoutOptions = () => ({ layoutExtras: '{}', stateVector: main.encodeStateVector() });
  const booted = await client.bootstrap(main.residentWorkerSnapshot()!, '', layoutOptions());

  const sessionId = main.paragraphIdentities().sessionId;
  const proposed = main.proposeChanges({
    expectVersion: main.version(),
    proposals: [0, 90, 200].map((index) => ({
      id: `p${index}`,
      paragraph: { kind: 'session', sessionId, story: 'body', paraId: ids[index]! },
      suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
      op: 'replaceText',
      search: 'dolor',
      replaceWith: 'a replacement long enough to rewrap the paragraph',
      occurrence: 'all',
    })),
  });
  expect(proposed.ok).toBe(true);
  main.adoptResidentWorkerLayout!(LAYOUT);
  const synced = await client.sync(
    main.residentWorkerSnapshot({
      knownStateVector: client.remoteStateVector(),
      knownFontsRevision: client.syncedFontsRevision(),
    })!,
    '',
    booted.caret.frameEpoch,
    false,
    layoutOptions()
  );
  const inWorker = JSON.parse(synced.layoutJson!) as { layout: { pages: unknown[] } };
  expect(inWorker.layout.pages.length).toBeGreaterThan(3);
  expect(inWorker).toEqual(JSON.parse(main.layoutDocumentWithRegionsRetainedJson(LAYOUT)));
});
