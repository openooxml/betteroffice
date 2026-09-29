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
import { createYrsSession, type YrsRenderEnv, type YrsSession } from './index';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(
  import.meta.dir,
  '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);
const SUGGEST = { author: 'Ann', date: '2026-09-29T12:00:00Z' };

type Preview = NonNullable<YrsRenderEnv['revisionPreview']>;

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

/** "Alpha beta gamma" with "beta" suggested as "BETA", and "Title" with a suggested "!". */
async function proposals(clientId: number) {
  const session = await createYrsSession({ clientId });
  sessions.push(session);
  const { paraId: first } = session.createStory('body', 'Alpha beta gammaTitle');
  const { secondParaId: second } = session.splitParagraph({
    story: 'body',
    paraId: first,
    offset: 16,
  });
  const search = (text: string, paraId: string) => ({
    kind: 'search' as const,
    text,
    within: { kind: 'paragraph' as const, story: 'body', paraId },
    view: 'accepted' as const,
  });
  const applied = session.applyEdits({
    expectVersion: session.version(),
    history: 'none',
    steps: [
      { op: 'replaceText', target: search('beta', first), text: 'BETA', suggest: SUGGEST },
      { op: 'insertText', target: search('Title', second), at: 'end', text: '!', suggest: SUGGEST },
    ],
  });
  if (!applied.ok) throw new Error(applied.failure.message);
  const [replace, insert] = applied.receipts.map((receipt) => receipt.revisionIds[0]);
  return { session, first, replace, insert };
}

interface LoweredRun {
  text?: string;
  pmStart: number;
  pmEnd: number;
  bold?: boolean;
  isInsertion?: boolean;
  isDeletion?: boolean;
}

interface LoweredParagraph {
  pmStart: number;
  pmEnd: number;
  runs: LoweredRun[];
}

function lower(session: YrsSession, revisionPreview?: Preview): LoweredParagraph[] {
  return session.yrsBlocksForStory('body', revisionPreview ? { revisionPreview } : {}) as LoweredParagraph[];
}

function runs(paragraph: LoweredParagraph): [string, number, number, string][] {
  return paragraph.runs.map((run) => [
    run.text ?? '',
    run.pmStart,
    run.pmEnd,
    run.isInsertion ? 'ins' : run.isDeletion ? 'del' : '',
  ]);
}

test('each decision renders its outcome at the source positions and reverts cleanly', async () => {
  const { session, replace } = await proposals(5201);
  const native = lower(session);
  expect(runs(native[0])).toEqual([
    ['Alpha ', 1, 7, ''],
    ['BETA', 7, 11, 'ins'],
    ['beta', 11, 15, 'del'],
    [' gamma', 15, 21, ''],
  ]);
  const expected = {
    proposed: runs(native[0]),
    accepted: [
      ['Alpha BETA', 1, 11, ''],
      [' gamma', 15, 21, ''],
    ],
    rejected: [
      ['Alpha ', 1, 7, ''],
      ['beta gamma', 11, 21, ''],
    ],
  };
  for (const state of ['accepted', 'rejected', 'proposed', 'rejected', 'accepted', 'proposed'] as const) {
    const blocks = lower(session, state === 'proposed' ? undefined : { [replace]: state });
    expect(runs(blocks[0])).toEqual(expected[state] as never);
    expect(blocks.map(({ pmStart, pmEnd }) => [pmStart, pmEnd])).toEqual(
      native.map(({ pmStart, pmEnd }) => [pmStart, pmEnd])
    );
    expect(blocks[1]).toEqual(native[1]);
  }
  expect(lower(session, {})).toEqual(native);
  expect(lower(session, { [replace]: 'proposed' } as unknown as Preview)).toEqual(native);
});

test('two proposals decide independently', async () => {
  const { session, replace, insert } = await proposals(5202);
  const native = lower(session);
  const mixed = lower(session, { [replace]: 'rejected', [insert]: 'accepted' });
  expect(runs(mixed[0])).toEqual([
    ['Alpha ', 1, 7, ''],
    ['beta gamma', 11, 21, ''],
  ]);
  expect(runs(mixed[1])).toEqual([['Title!', 23, 29, '']]);
  const second = lower(session, { [insert]: 'rejected' });
  expect(second[0]).toEqual(native[0]);
  expect(runs(second[1])).toEqual([['Title', 23, 28, '']]);
});

test('a decided run keeps its own formatting and only loses its markup', async () => {
  const session = await createYrsSession({ clientId: 5203 });
  sessions.push(session);
  const { paraId } = session.createStory('body', 'plain bold tail');
  session.applyRawOps('body', [{ op: 'format', index: 6, len: 4, attrs: { bold: true } }]);
  const applied = session.applyEdits({
    expectVersion: session.version(),
    history: 'none',
    steps: [
      {
        op: 'replaceText',
        target: {
          kind: 'search',
          text: 'bold',
          within: { kind: 'paragraph', story: 'body', paraId },
          view: 'accepted',
        },
        text: 'BOLD',
        suggest: SUGGEST,
      },
    ],
  });
  if (!applied.ok) throw new Error(applied.failure.message);
  const [revision] = applied.receipts[0].revisionIds;
  const bolded = (blocks: LoweredParagraph[]) =>
    blocks[0].runs.filter((run) => run.bold).map((run) => [run.text, !!run.isInsertion, !!run.isDeletion]);
  expect(bolded(lower(session))).toEqual([
    ['BOLD', true, false],
    ['bold', false, true],
  ]);
  expect(bolded(lower(session, { [revision]: 'accepted' }))).toEqual([['BOLD', false, false]]);
  expect(bolded(lower(session, { [revision]: 'rejected' }))).toEqual([['bold', false, false]]);
  expect(bolded(lower(session))).toEqual([
    ['BOLD', true, false],
    ['bold', false, true],
  ]);
});

test('previewing leaves the document, its revisions and undo history untouched', async () => {
  const { session, first, replace, insert } = await proposals(5204);
  session.registerFont(new Uint8Array(readFileSync(FONT)));
  session.insertText({ story: 'body', paraId: first, offset: 0 }, '>');
  expect(session.undo()).toBe(true);
  const state = session.encodeState();
  const version = session.version();
  const revisions = session.listRevisions();
  const history = [session.canUndo(), session.canRedo()];
  for (const decision of ['accepted', 'rejected'] as const) {
    const revisionPreview = { [replace]: decision, [insert]: decision };
    lower(session, revisionPreview);
    session.layoutDocumentWithRegionsJson(layoutInput({ revisionPreview }));
  }
  expect(session.encodeState()).toEqual(state);
  expect(session.version()).toBe(version);
  expect(session.listRevisions()).toEqual(revisions);
  expect([session.canUndo(), session.canRedo()]).toEqual(history);
  expect(session.redo()).toBe(true);
  expect(session.paragraphs('body')[0].text.startsWith('>Alpha')).toBe(true);
  expect(session.undo()).toBe(true);
  expect(session.paragraphs('body')[0].text.startsWith('Alpha')).toBe(true);
  expect(session.listRevisions()).toEqual(revisions);
});

function layoutInput(renderEnv: YrsRenderEnv): string {
  return JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
    renderEnv,
  });
}

/** Painted text, split where the revision kind changes. */
function painted(frame: RetainedFrame): [string, string][] {
  const segments: [string, string][] = [];
  for (const primitive of frame.displayList.pages.flatMap((page) => page.primitives)) {
    if (primitive.kind !== 'text' && primitive.kind !== 'glyphRun') continue;
    const kind = primitive.revision?.kind ?? '';
    const last = segments.at(-1);
    if (last && last[1] === kind) last[0] += primitive.text;
    else segments.push([primitive.text, kind]);
  }
  return segments;
}

test('worker frames match the main thread, and a preview change retires older frames', async () => {
  const { session: main, first, replace, insert } = await proposals(5205);
  main.registerFont(new Uint8Array(readFileSync(FONT)));
  const accepted = { [replace]: 'accepted', [insert]: 'accepted' } as const;
  main.yrsBlocksForStory('body', { revisionPreview: accepted });
  main.layoutDocumentWithRegionsJson(layoutInput({ revisionPreview: accepted }));
  main.setSelection({ story: 'body', paraId: first, offset: 0 });
  let local = applyFrameDeltaOwned(null, decodeFrameDelta(main.buildDisplayListFrame('{}', 0)));
  expect(painted(local)).toEqual([['Alpha BETA gammaTitle!', '']]);

  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  const booted = await client.bootstrap(main.residentWorkerSnapshot()!, '{}');
  const bootFrame = booted.frame.slice();
  let remote = applyFrameDeltaOwned(null, decodeFrameDelta(booted.frame));
  expect(painted(remote)).toEqual(painted(local));

  const rejected = { [replace]: 'rejected' } as const;
  main.layoutDocumentWithRegionsJson(layoutInput({ revisionPreview: rejected }));
  local = applyFrameDeltaOwned(
    local,
    decodeFrameDelta(main.buildDisplayListFrame('{}', local.frameEpoch))
  );
  expect(painted(local)).toEqual([
    ['Alpha beta gammaTitle', ''],
    ['!', 'ins'],
  ]);
  const synced = await client.sync(
    main.residentWorkerSnapshot({
      knownStateVector: client.remoteStateVector(),
      knownFontsRevision: client.syncedFontsRevision(),
    })!,
    '{}',
    remote.frameEpoch
  );
  remote = applyFrameDeltaOwned(remote, decodeFrameDelta(synced.frame));
  expect(painted(remote)).toEqual(painted(local));

  let stale: unknown = null;
  try {
    applyFrameDeltaOwned(remote, decodeFrameDelta(bootFrame));
  } catch (error) {
    stale = error;
  }
  expect(stale).toBeInstanceOf(FrameDeltaError);
  expect((stale as FrameDeltaError).code).toBe('stale-frame');
});
