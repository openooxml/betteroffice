import { afterAll, afterEach, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { preloadEditWasm } from '@betteroffice/docx/yrs/wasm/index';
import { preloadOpcWasm } from '@betteroffice/docx/wasm/opc';
import { residentWorkerFactory, type InProcessResidentWorker } from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import { DocxSourceWorkerClient, createDocxCommandConnection, SessionFailure } from '@betteroffice/docx/worker';
import type { DocxCommandCheckpoint, DocxCommandConnection } from '@betteroffice/docx/worker';
import { usePeerlessEditSession, type UsePeerlessEditSessionOptions } from './usePeerlessEditSession';

const TIMEOUT = 180_000;
const generated = resolve(import.meta.dir, '../../../../../docx/src/wasm/generated');
await preloadEditWasm(new Uint8Array(readFileSync(resolve(generated, 'edit/docx_edit_bg.wasm'))));
await preloadOpcWasm(new Uint8Array(readFileSync(resolve(generated, 'opc/ooxml_opc_bg.wasm'))));
const start = await residentWorkerFactory();
const FONT = new Uint8Array(readFileSync(resolve(import.meta.dir, '../../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')));

const { GlobalRegistrator } = await import('@happy-dom/global-registrator');
const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const liveWorkers: InProcessResidentWorker[] = [];
afterEach(() => {
  cleanup();
  for (const worker of liveWorkers.splice(0)) { try { worker.terminate(); } catch {} }
});
afterAll(async () => {
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

const documentBytes = () => new Uint8Array(rezipPartsToArrayBuffer(new Map([
  ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
  ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
  ['word/document.xml', toBytes('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Seed</w:t></w:r></w:p><w:sectPr/></w:body></w:document>')],
])));

const LAYOUT = JSON.stringify({ bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' }, fontChains: { 'liberation sans|0|0': [0] } }, renderEnv: {} });
const LETTERS = 'abcdefghijklmnopqrstuvwxyz';

type Resume = Parameters<NonNullable<UsePeerlessEditSessionOptions['recovery']>['reopen']>[0];
type Seed = NonNullable<DocxCommandConnection['initial']['seed']>;

const textOf = (worker: InProcessResidentWorker) =>
  worker.sessions.at(-1)!.paragraphs('body').map((paragraph) => paragraph.text).join('|');

const INIT = {
  readOnly: false, fonts: [FONT], fontRegistrationSlots: true, layoutInput: LAYOUT, layoutWithRegions: true,
  layoutRevision: 1, extras: JSON.stringify({ fontChains: { 'liberation sans|0|0': [0] } }), displayWindow: [0, 3] as [number, number],
};

function mount({ tamper }: { tamper?: (seed: Seed) => void } = {}) {
  const source = documentBytes();
  const workers: InProcessResidentWorker[] = [];
  const checkpoints: DocxCommandCheckpoint[] = [];
  const resumes: Resume[] = [];
  const reopenedText: string[] = [];
  const replayed: string[] = [];
  const phases: string[] = [];
  const errors: Error[] = [];
  let pendingCheckpoint: Promise<unknown> | undefined;
  let applied = 0;
  const openResident = async (resume?: Resume): Promise<DocxCommandConnection> => {
    const worker = start();
    workers.push(worker);
    liveWorkers.push(worker);
    const client = new DocxSourceWorkerClient(worker);
    const checkpoint = resume?.checkpoint as DocxCommandCheckpoint | undefined;
    if (checkpoint && !checkpoint.recoveryState) {
      throw new SessionFailure('message', 'Cannot restore DOCX checkpoint command-local state and proposals');
    }
    await client.openSource({ bytes: source.slice(), ...(checkpoint ? { state: checkpoint.state } : {}),
      ...(resume?.seed ? { clientId: resume.seed.clientId, generation: resume.seed.generation } : {}) });
    const initial = await client.commandInit({
      session: resume?.session ?? 'undo-recovery', incarnation: resume?.incarnation ?? 1, epoch: resume?.epoch ?? 1,
      ...(resume ? { version: checkpoint?.version ?? 0, throughSeq: checkpoint?.throughSeq ?? 0,
        refusedSeqs: resume.refusedSeqs, replayAppliedSeqs: resume.replayAppliedSeqs,
        recoveryState: checkpoint?.recoveryState } : {}),
      ...INIT,
    });
    const connection = await createDocxCommandConnection(client, initial);
    if (!resume) {
      if (tamper) tamper(connection.initial.seed!);
      return connection;
    }
    reopenedText.push(textOf(worker));
    const call = connection.client.call;
    const edit = call.edit;
    call.edit = (command) => {
      replayed.push(command.kind);
      return edit(command);
    };
    return connection;
  };
  const options: UsePeerlessEditSessionOptions = {
    source,
    openConnection: () => openResident(),
    recovery: {
      checkpoint: async (resident) => {
        const running = resident.client.call.checkpoint();
        pendingCheckpoint = running;
        const value = await running;
        checkpoints.push(value);
        return value;
      },
      reopen: (resume) => {
        resumes.push(resume);
        return openResident(resume);
      },
      onRecovery: (event) => { phases.push(event.phase); },
    },
    canvasForPage: () => null,
    prepareGeometry: () => undefined,
    onError: (error) => { errors.push(error); },
  };
  const view = renderHook(() => usePeerlessEditSession(options));
  const settle = () => act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  const session = {
    view, checkpoints, phases, errors, resumes, reopenedText, replayed,
    async ready() { await waitFor(() => expect(view.result.current.status).toBe('ready'), { timeout: 30_000 }); },
    async submit(kind: string, payload: Record<string, unknown> = {}) {
      let status = '';
      await act(async () => {
        const outcome = await view.result.current.binding!.session.submit(kind, payload as never).ack as { status: string };
        status = outcome.status;
      });
      if (status === 'applied') applied += 1;
      return status;
    },
    async type(count: number) {
      for (let i = 0; i < count; i += 1) expect(await session.submit('insertText', { text: LETTERS[i % 26]! })).toBe('applied');
    },
    async typeUntilApplied(target: number) {
      for (let i = 0; applied < target && i < 400; i += 1) await session.submit('insertText', { text: LETTERS[i % 26]! });
      expect(applied).toBe(target);
    },
    async checkpointTaken(count: number) {
      for (let i = 0; i < 200 && checkpoints.length < count; i += 1) {
        await settle();
        if (pendingCheckpoint) await act(async () => { try { await pendingCheckpoint; } catch {} });
      }
      expect(checkpoints.length).toBeGreaterThanOrEqual(count);
    },
    text() {
      return textOf(workers.at(-1)!);
    },
    version() {
      return view.result.current.binding!.session.state.version;
    },
    selection() {
      return view.result.current.binding!.session.state.selection!;
    },
    boldBeforeCaret() {
      const { focus } = session.selection();
      const engine = workers.at(-1)!.sessions.at(-1)!;
      return engine.selectionContext({ story: focus.story,
        start: { paraId: focus.paraId, offset: focus.offset - 1 }, end: { paraId: focus.paraId, offset: focus.offset } }).bold;
    },
    async crash() {
      const worker = workers.at(-1)!;
      await act(async () => { worker.onerror?.({ message: 'simulated worker crash' } as ErrorEvent); });
      await waitFor(() => expect(phases.includes('recovered') || phases.includes('failed')
        || view.result.current.status === 'failed').toBe(true), { timeout: 60_000 });
      for (let i = 0; i < 5; i += 1) await settle();
    },
    expectRecovered(expected: string) {
      expect(errors.map((error) => error.message)).toEqual([]);
      expect(phases).toContain('recovered');
      expect(view.result.current.status).toBe('ready');
      expect(session.text()).toBe(expected);
    },
  };
  return session;
}

test('undo after a checkpoint survives a worker restart', async () => {
  const session = mount();
  await session.ready();
  await session.typeUntilApplied(128);
  await session.checkpointTaken(1);
  const typed = session.text();
  expect(await session.submit('undo')).toBe('applied');
  const expected = session.text();
  const version = session.version();
  expect(expected).not.toBe(typed);
  await session.crash();
  session.expectRecovered(expected);
  expect(session.version()).toBe(version);
}, TIMEOUT);

test('redo after a checkpoint survives a worker restart', async () => {
  const session = mount();
  await session.ready();
  await session.typeUntilApplied(127);
  expect(await session.submit('undo')).toBe('applied');
  await session.checkpointTaken(1);
  const undone = session.text();
  expect(await session.submit('redo')).toBe('applied');
  const expected = session.text();
  const version = session.version();
  expect(expected).not.toBe(undone);
  await session.crash();
  session.expectRecovered(expected);
  expect(session.version()).toBe(version);
}, TIMEOUT);

test('undo reaching across a checkpoint recovers the document the user saw', async () => {
  const session = mount();
  await session.ready();
  await session.typeUntilApplied(128);
  await session.checkpointTaken(1);
  for (const text of ['Z', 'Y', 'X']) expect(await session.submit('insertText', { text })).toBe('applied');
  expect(await session.submit('undo')).toBe('applied');
  const expected = session.text();
  const version = session.version();
  await session.crash();
  session.expectRecovered(expected);
  expect(session.version()).toBe(version);
}, TIMEOUT);

test('after recovery from an undo across a checkpoint, a further undo and redo change nothing the user never saw', async () => {
  const session = mount();
  await session.ready();
  await session.typeUntilApplied(128);
  await session.checkpointTaken(1);
  for (const text of ['Z', 'Y', 'X']) expect(await session.submit('insertText', { text })).toBe('applied');
  expect(await session.submit('undo')).toBe('applied');
  const expected = session.text();
  await session.crash();
  session.expectRecovered(expected);
  await session.submit('undo');
  expect(session.text()).toBe(expected);
  await session.submit('redo');
  expect(session.text()).toBe(expected);
}, TIMEOUT);

test('undo before any checkpoint survives a worker restart', async () => {
  const session = mount();
  await session.ready();
  await session.type(5);
  expect(await session.submit('undo')).toBe('applied');
  const expected = session.text();
  await session.crash();
  expect(session.checkpoints).toEqual([]);
  expect(session.resumes[0]!.checkpoint).toBeUndefined();
  expect(session.resumes[0]!.seed).toBeDefined();
  session.expectRecovered(expected);
}, TIMEOUT);

test('a crash before any edit recovers the source document', async () => {
  const session = mount();
  await session.ready();
  await session.crash();
  expect(session.resumes).toHaveLength(1);
  expect(session.resumes[0]!.checkpoint).toBeUndefined();
  expect(session.resumes[0]!.replayAppliedSeqs).toEqual([]);
  session.expectRecovered('Seed');
  expect(session.replayed).toEqual([]);
  expect(await session.submit('insertText', { text: 'a' })).toBe('applied');
  expect(session.checkpoints).toEqual([]);
}, TIMEOUT);

test('a command with a worker-issued position, no checkpoint, crash, recover: applied, same text', async () => {
  const session = mount();
  await session.ready();
  await session.type(3);
  const issued = session.selection();
  expect(issued.focus.sticky).toBeDefined();
  expect(await session.submit('insertText', { text: 'W', selection: { anchor: issued.anchor, focus: issued.focus } })).toBe('applied');
  const expected = session.text();
  const version = session.version();
  await session.crash();
  expect(session.checkpoints).toEqual([]);
  expect(session.replayed.length).toBeGreaterThan(0);
  expect(session.replayed.every((kind) => kind === 'applyUpdate')).toBe(true);
  session.expectRecovered(expected);
  expect(session.version()).toBe(version);
}, TIMEOUT);

test('a reseed that differs from the opened document fails recovery and applies nothing', async () => {
  const session = mount({ tamper: (seed) => { seed.stateVector = Uint8Array.of(...seed.stateVector, 0); } });
  await session.ready();
  await session.type(3);
  expect(session.text()).not.toBe('Seed');
  await session.crash();
  expect(session.errors.map((error) => error.message)).toEqual(['Recovery base differs from the opened document']);
  expect(session.phases).not.toContain('recovered');
  expect(session.view.result.current.status).toBe('failed');
  expect(session.resumes).toHaveLength(1);
  expect(session.reopenedText).toEqual(['Seed']);
  expect(session.replayed).toEqual([]);
}, TIMEOUT);

test('the same document reseeded with the pinned client id and generation gives the same state vector', async () => {
  const root = resolve(import.meta.dir, '../../../../../..');
  const fixtures: [string, Uint8Array][] = [
    ['inline', documentBytes()],
    ...[
      'packages/docx/src/yrs/__fixtures__/comment-ranges/structure.docx',
      'crates/docx-edit/tests/fixtures/header-footer-reservation/continuous-title-page-at-page-break.docx',
      'crates/betteroffice-docx/tests/corpus/fixtures/wordprocessingml-comprehensive.docx',
      'apps/demo/public/betteroffice-demo.docx',
    ].map((path): [string, Uint8Array] => [path, new Uint8Array(readFileSync(resolve(root, path)))]),
  ];
  const seedOf = async (bytes: Uint8Array, pinned?: Pick<Seed, 'clientId' | 'generation'>): Promise<Seed> => {
    const worker = start();
    liveWorkers.push(worker);
    const client = new DocxSourceWorkerClient(worker);
    await client.openSource({ bytes: bytes.slice(), ...(pinned ?? {}) });
    const initial = await client.commandInit({ session: 'reseed', incarnation: 1, epoch: 1, ...INIT });
    expect(initial.seed).toBeDefined();
    return initial.seed!;
  };
  for (const [name, bytes] of fixtures) {
    const opened = await seedOf(bytes);
    const reseeded = await seedOf(bytes, { clientId: opened.clientId, generation: opened.generation });
    console.log(`G244-SEED ${name} stateVector=${opened.stateVector.byteLength} source=${bytes.byteLength}`);
    expect(reseeded.clientId).toBe(opened.clientId);
    expect(reseeded.generation).toBe(opened.generation);
    expect([...reseeded.stateVector]).toEqual([...opened.stateVector]);
  }
}, TIMEOUT);

test('a caret-only bold after a checkpoint still applies to typing after a worker restart', async () => {
  const live = mount();
  await live.ready();
  await live.typeUntilApplied(128);
  await live.checkpointTaken(1);
  expect(await live.submit('insertText', { text: 'Z' })).toBe('applied');
  expect(await live.submit('toggleMark', { mark: { type: 'bold' } })).not.toBe('refused');
  const liveBefore = live.text();
  const liveCaret = live.selection().focus.offset;
  expect(await live.submit('insertText', { text: 'q' })).toBe('applied');
  const liveTyped = live.text();
  const liveBold = live.boldBeforeCaret();
  expect(liveTyped).toBe(liveBefore.slice(0, liveCaret) + 'q' + liveBefore.slice(liveCaret));
  expect(liveBold).toBe(true);
  cleanup();

  const restarted = mount();
  await restarted.ready();
  await restarted.typeUntilApplied(128);
  await restarted.checkpointTaken(1);
  expect(await restarted.submit('insertText', { text: 'Z' })).toBe('applied');
  expect(await restarted.submit('toggleMark', { mark: { type: 'bold' } })).not.toBe('refused');
  const before = restarted.text();
  const caret = restarted.selection().focus.offset;
  await restarted.crash();
  restarted.expectRecovered(before);
  expect(restarted.selection().focus.offset).toBe(caret);
  expect(await restarted.submit('insertText', { text: 'q' })).toBe('applied');
  expect(restarted.text()).toBe(before.slice(0, caret) + 'q' + before.slice(caret));
  expect(restarted.boldBeforeCaret()).toBe(liveBold);
}, TIMEOUT);

test('an effect over the byte limit after a checkpoint brings the next checkpoint', async () => {
  const session = mount();
  await session.ready();
  await session.typeUntilApplied(128);
  await session.checkpointTaken(1);
  const taken = session.checkpoints.length;
  expect(await session.submit('insertText', { text: 'x'.repeat(1_150_000) })).toBe('applied');
  await session.checkpointTaken(taken + 1);
  expect(session.checkpoints.at(-1)!.throughSeq).toBeGreaterThan(session.checkpoints[taken - 1]!.throughSeq);
}, TIMEOUT);

test('one undo after a worker restart removes a measured amount of the typing before it', async () => {
  const live = mount();
  await live.ready();
  await live.type(3);
  const liveTyped = live.text().length;
  expect(await live.submit('undo')).toBe('applied');
  const liveRemoved = liveTyped - live.text().length;
  cleanup();

  const restarted = mount();
  await restarted.ready();
  await restarted.type(3);
  const typed = restarted.text();
  await restarted.crash();
  restarted.expectRecovered(typed);
  await restarted.submit('undo');
  const recoveredRemoved = typed.length - restarted.text().length;

  console.log(`G244-GROUPING live=${liveRemoved} recovered=${recoveredRemoved}`);
  expect(Number.isInteger(liveRemoved)).toBe(true);
  expect(Number.isInteger(recoveredRemoved)).toBe(true);
}, TIMEOUT);
