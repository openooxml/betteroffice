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
