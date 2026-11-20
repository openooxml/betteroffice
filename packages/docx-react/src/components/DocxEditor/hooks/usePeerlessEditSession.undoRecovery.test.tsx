import { afterAll, afterEach, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { preloadEditWasm } from '@betteroffice/docx/yrs/wasm/index';
import { preloadOpcWasm } from '@betteroffice/docx/wasm/opc';
import { residentWorkerFactory, type InProcessResidentWorker } from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import { createYrsSession } from '@betteroffice/docx/yrs';
import { DocxSourceWorkerClient, createDocxCommandConnection, createDocxWorkerEditSession, SessionFailure } from '@betteroffice/docx/worker';
import type { DocxCommandCheckpoint, DocxCommandConnection } from '@betteroffice/docx/worker';
import type { Document } from '@betteroffice/docx/types/document';
import type { CommandTicket } from '../../../../../../shared/office-session/commands/types';
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

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const documentBytes = (text = 'Seed', header?: string) => new Uint8Array(rezipPartsToArrayBuffer(new Map<string, Uint8Array>([
  ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    (header === undefined ? '' : '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>') + '</Types>')],
  ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
  ...(header === undefined ? [] : [
    ['word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH1" Type="${REL}/header" Target="header1.xml"/></Relationships>`)],
    ['word/header1.xml', toBytes(`<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>${header}</w:t></w:r></w:p></w:hdr>`)],
  ] as [string, Uint8Array][]),
  ['word/document.xml', toBytes(header === undefined
    ? `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
    : `<w:document xmlns:w="${W}" xmlns:r="${REL}"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/></w:sectPr></w:body></w:document>`)],
])));

const LAYOUT = JSON.stringify({ bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' }, fontChains: { 'liberation sans|0|0': [0] } }, renderEnv: {} });
const LETTERS = 'abcdefghijklmnopqrstuvwxyz';

type Resume = Parameters<NonNullable<UsePeerlessEditSessionOptions['recovery']>['reopen']>[0];
type Seed = NonNullable<DocxCommandConnection['initial']['seed']>;
type RecoveryReport = Parameters<NonNullable<NonNullable<UsePeerlessEditSessionOptions['recovery']>['onRecovery']>>[0];

const textOf = (worker: InProcessResidentWorker) =>
  worker.sessions.at(-1)!.paragraphs('body').map((paragraph) => paragraph.text).join('|');

const INIT = {
  readOnly: false, fonts: [FONT], fontRegistrationSlots: true, layoutInput: LAYOUT, layoutWithRegions: true,
  layoutRevision: 1, extras: JSON.stringify({ fontChains: { 'liberation sans|0|0': [0] } }), displayWindow: [0, 3] as [number, number],
};

interface MountOptions {
  tamper?: (seed: Seed) => void;
  source?: Uint8Array;
  reopenSource?: Uint8Array;
  collaboration?: boolean;
  gateCheckpoint?: boolean;
}

function mount({ tamper, source = documentBytes(), reopenSource, collaboration, gateCheckpoint }: MountOptions = {}) {
  const workers: InProcessResidentWorker[] = [];
  const clients: DocxSourceWorkerClient[] = [];
  let openGate = () => {};
  const gate = { requested: 0, opened: new Promise<void>((resolve) => { openGate = resolve; }) };
  const checkpoints: DocxCommandCheckpoint[] = [];
  const resumes: Resume[] = [];
  const reopenedText: string[] = [];
  const replayed: string[] = [];
  const phases: string[] = [];
  const events: RecoveryReport[] = [];
  const errors: Error[] = [];
  let pendingCheckpoint: Promise<unknown> | undefined;
  let applied = 0;
  const openResident = async (resume?: Resume): Promise<DocxCommandConnection> => {
    const worker = start();
    workers.push(worker);
    liveWorkers.push(worker);
    const client = new DocxSourceWorkerClient(worker);
    clients.push(client);
    const checkpoint = resume?.checkpoint as DocxCommandCheckpoint | undefined;
    if (checkpoint && !checkpoint.recoveryState) {
      throw new SessionFailure('message', 'Cannot restore DOCX checkpoint command-local state and proposals');
    }
    await client.openSource({ bytes: (resume && reopenSource ? reopenSource : source).slice(), ...(checkpoint ? { state: checkpoint.state } : {}),
      ...(resume?.seed ? { clientId: resume.seed.clientId, generation: resume.seed.generation } : {}),
      ...(collaboration ? { collaboration: { clientId: 9100 + workers.length } } : {}) });
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
        if (gateCheckpoint) {
          gate.requested += 1;
          await gate.opened;
        }
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
      onRecovery: (event) => { phases.push(event.phase); events.push(event); },
    },
    canvasForPage: () => null,
    prepareGeometry: () => undefined,
    onError: (error) => { errors.push(error); },
  };
  const view = renderHook(() => usePeerlessEditSession(options));
  const settle = () => act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  const session = {
    view, checkpoints, phases, events, errors, resumes, reopenedText, replayed, settle,
    gate: { get requested() { return gate.requested; }, open: () => openGate() },
    worker() { return workers.at(-1)!; },
    client() { return clients.at(-1)!; },
    engine() { return workers.at(-1)!.sessions.at(-1)!; },
    stateVector() { return [...workers.at(-1)!.sessions.at(-1)!.encodeStateVector()].join(','); },
    async outcome(status: Promise<string>, timeout = 30_000) {
      let timer: ReturnType<typeof setTimeout> | undefined;
      let settled = 'unsettled';
      await act(async () => {
        settled = await Promise.race([status, new Promise<string>((resolve) => { timer = setTimeout(() => resolve('unsettled'), timeout); })]);
      });
      clearTimeout(timer);
      return settled;
    },
    activePart() { return view.result.current.binding!.session.state.editState?.activePart ?? null; },
    async save() {
      await act(async () => { await clients.at(-1)!.save({ comments: [] }); });
    },
    async submitQueued(kind: string, payload: Record<string, unknown> = {}, after?: CommandTicket): Promise<{ status: Promise<string>; detail: Promise<string>; ticket: CommandTicket }> {
      let status!: Promise<string>;
      let detail!: Promise<string>;
      let ticket!: CommandTicket;
      await act(async () => {
        ticket = view.result.current.binding!.session.submit(kind, payload as never, after ? { after } : {});
        const ack = ticket.ack as Promise<{ status: string; reason?: string }>;
        status = ack.then((outcome) => outcome.status, (error: Error) => `rejected: ${error.message}`);
        detail = ack.then((outcome) => JSON.stringify({ status: outcome.status, reason: outcome.reason }),
          (error: Error) => `rejected: ${error.message}`);
      });
      return { status, detail, ticket };
    },
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
  expect(recoveredRemoved).toBe(0);
}, TIMEOUT);

test('a save between acknowledged edits does not leave the recovered document behind', async () => {
  const session = mount();
  await session.ready();
  expect(await session.submit('paragraphBreak')).toBe('applied');
  await session.save();
  expect(await session.submit('insertText', { text: 'B' })).toBe('applied');
  const expected = session.text();
  const version = session.version();
  expect(expected.split('|')).toHaveLength(2);
  expect(expected).toContain('B');
  await session.crash();
  expect(session.checkpoints).toEqual([]);
  session.expectRecovered(expected);
  expect(session.version()).toBe(version);
  expect(session.engine().hasPendingUpdate()).toBe(false);
}, TIMEOUT);

test('a save between two acknowledged edits does not leave the recovered document behind', async () => {
  const session = mount();
  await session.ready();
  expect(await session.submit('insertText', { text: 'A' })).toBe('applied');
  await session.save();
  expect(await session.submit('insertText', { text: 'B' })).toBe('applied');
  const expected = session.text();
  const version = session.version();
  await session.crash();
  expect(session.checkpoints).toEqual([]);
  session.expectRecovered(expected);
  expect(session.version()).toBe(version);
  expect(session.engine().hasPendingUpdate()).toBe(false);
}, TIMEOUT);

test('a save that publishes paragraph ids between acknowledged edits recovers the document', async () => {
  const source = documentBytes('Seed');
  const session = mount({ source });
  await session.ready();
  expect(await session.submit('paragraphBreak')).toBe('applied');
  const before = session.stateVector();
  await session.save();
  if (session.stateVector() === before) throw new Error('save wrote nothing; mechanism not reached');
  expect(await session.submit('insertText', { text: 'B' })).toBe('applied');
  const expected = session.text();
  const version = session.version();
  expect(expected.split('|')).toHaveLength(2);
  expect(expected).toContain('B');
  await session.crash();
  expect(session.checkpoints).toEqual([]);
  session.expectRecovered(expected);
  expect(session.version()).toBe(version);
  expect(session.engine().hasPendingUpdate()).toBe(false);
}, TIMEOUT);

test('a save landing while a command awaits keeps stream order', async () => {
  const session = mount();
  await session.ready();
  expect(await session.submit('paragraphBreak')).toBe('applied');
  const worker = session.worker();
  worker.hold();
  const typed = await session.submitQueued('insertText', { text: 'B' });
  await waitFor(() => expect(textOf(worker)).toContain('B'), { timeout: 30_000 });
  let saved = false;
  let saving!: Promise<unknown>;
  await act(async () => { saving = session.client().save({ comments: [] }).then(() => { saved = true; }); });
  await waitFor(() => {
    worker.release((reply) => 'saved' in reply);
    expect(saved).toBe(true);
  }, { timeout: 30_000 });
  worker.release();
  await act(async () => { await saving; });
  expect(await session.outcome(typed.status)).toBe('applied');
  expect(await session.submit('insertText', { text: 'C' })).toBe('applied');
  const expected = session.text();
  const version = session.version();
  expect(expected).toContain('BC');
  await session.crash();
  expect(session.checkpoints).toEqual([]);
  session.expectRecovered(expected);
  expect(session.version()).toBe(version);
  expect(session.engine().hasPendingUpdate()).toBe(false);
}, TIMEOUT);

test('a frame-time update between two commands is in the log', async () => {
  const session = mount();
  await session.ready();
  expect(await session.submit('insertText', { text: 'A' })).toBe('applied');
  const engine = session.engine();
  const remote = await createYrsSession({ clientId: 9401 });
  try {
    remote.loadState(engine.encodeState());
    const paragraph = remote.paragraphs('body')[0]!;
    remote.insertText({ story: 'body', paraId: paragraph.paraId, offset: paragraph.text.length }, 'M');
    const update = remote.encodeStateAsUpdate(engine.encodeStateVector());
    await act(async () => { await session.client().applyRemoteUpdate(update); });
  } finally { remote.destroy(); }
  expect(session.text()).toContain('M');
  expect(await session.submit('insertText', { text: 'Z' })).toBe('applied');
  const expected = session.text();
  const version = session.version();
  for (const letter of ['A', 'M', 'Z']) expect(expected).toContain(letter);
  await session.crash();
  expect(session.checkpoints).toEqual([]);
  session.expectRecovered(expected);
  expect(session.version()).toBe(version);
  expect(session.engine().hasPendingUpdate()).toBe(false);
}, TIMEOUT);

test('idle saves do not grow the log without bound', async () => {
  const session = mount();
  await session.ready();
  const before = session.stateVector();
  await session.save();
  const published = session.stateVector() !== before;
  for (let i = 1; i < 300; i += 1) await session.save();
  await session.settle();
  const expected = session.text();
  await session.crash();
  console.log(`G244-IDLE-SAVES published=${published} checkpoints=${session.checkpoints.length} replayed=${session.replayed.length}`);
  session.expectRecovered(expected);
  expect(session.engine().hasPendingUpdate()).toBe(false);
  expect(session.checkpoints.length > 0 || session.replayed.length <= 128).toBe(true);
}, 600_000);

test('an undo still in flight at a crash before the first checkpoint is refused and reported', async () => {
  const session = mount();
  await session.ready();
  const before = session.text();
  expect(await session.submit('insertText', { text: 'A' })).toBe('applied');
  const typed = session.text();
  expect(typed).not.toBe(before);
  expect(typed).toContain('A');
  const worker = session.worker();
  worker.hold();
  const undo = await session.submitQueued('undo');
  await waitFor(() => expect(textOf(worker)).toBe(before), { timeout: 30_000 });
  await session.crash();
  expect(session.checkpoints).toEqual([]);
  session.expectRecovered(typed);
  const status = await session.outcome(undo.status);
  console.log(`G244-INFLIGHT-UNDO status=${status} detail=${await session.outcome(undo.detail)} phases=${JSON.stringify(session.phases)}`);
  expect(['refused', 'cancelled']).toContain(status);
  expect(session.phases).toContain('diverged');
  expect(session.events.some((event) => event.phase === 'diverged' && event.diverged.length > 0 && event.historyReset)).toBe(true);
  expect(session.engine().hasPendingUpdate()).toBe(false);
}, TIMEOUT);

test('a redo still in flight at a crash before the first checkpoint is refused and reported', async () => {
  const session = mount();
  await session.ready();
  const before = session.text();
  expect(await session.submit('insertText', { text: 'A' })).toBe('applied');
  const typed = session.text();
  expect(await session.submit('undo')).toBe('applied');
  expect(session.text()).toBe(before);
  const worker = session.worker();
  worker.hold();
  const redo = await session.submitQueued('redo');
  await waitFor(() => expect(textOf(worker)).toBe(typed), { timeout: 30_000 });
  await session.crash();
  expect(session.checkpoints).toEqual([]);
  session.expectRecovered(before);
  const status = await session.outcome(redo.status);
  console.log(`G244-INFLIGHT-REDO status=${status} detail=${await session.outcome(redo.detail)} phases=${JSON.stringify(session.phases)}`);
  expect(['refused', 'cancelled']).toContain(status);
  expect(session.phases).toContain('diverged');
  expect(session.events.some((event) => event.phase === 'diverged' && event.diverged.length > 0 && event.historyReset)).toBe(true);
  expect(session.engine().hasPendingUpdate()).toBe(false);
}, TIMEOUT);

test('with collaboration the logged entries keep their original commands', async () => {
  const session = mount({ collaboration: true });
  await session.ready();
  await session.typeUntilApplied(128);
  await session.checkpointTaken(1);
  expect(await session.submit('insertText', { text: 'Z' })).toBe('applied');
  const expected = session.text();
  await session.crash();
  expect(session.resumes[0]!.seed).toBeUndefined();
  expect(session.replayed).toContain('insertText');
  expect(session.replayed).not.toContain('applyUpdate');
  session.expectRecovered(expected);
}, TIMEOUT);

test('a same-length different source is refused', async () => {
  const session = mount({ reopenSource: documentBytes('Reed') });
  await session.ready();
  await session.type(1);
  await session.crash();
  expect(session.errors.map((error) => error.message)).toEqual(['Recovery base differs from the opened document']);
  expect(session.phases).not.toContain('recovered');
  expect(session.view.result.current.status).toBe('failed');
  expect(session.resumes).toHaveLength(1);
  expect(session.reopenedText).toEqual(['Reed']);
  expect(session.replayed).toEqual([]);
}, TIMEOUT);

test('opening a header survives a worker restart before a checkpoint', async () => {
  const session = mount({ source: documentBytes('Seed', 'Head') });
  await session.ready();
  const expectVersion = session.engine().proposalEngine.version();
  expect(await session.submit('openPart', { kind: 'header', section: 'final', page: 0, expectVersion })).toBe('noop');
  expect(session.activePart()?.kind).toBe('header');
  const story = session.selection().focus.story;
  expect(story).toMatch(/^hf:/);
  await session.crash();
  expect(session.checkpoints).toEqual([]);
  session.expectRecovered('Seed');
  expect(session.activePart()?.kind).toBe('header');
  expect(session.selection().focus.story).toBe(story);
}, TIMEOUT);

test('typing queued across a checkpoint', async () => {
  const session = mount({ gateCheckpoint: true });
  await session.ready();
  await session.typeUntilApplied(127);
  expect(await session.submit('insertText', { text: 'P' })).toBe('applied');
  await waitFor(() => expect(session.gate.requested).toBe(1), { timeout: 30_000 });
  const first = await session.submitQueued('insertText', { text: 'Q' });
  const queued = [first, await session.submitQueued('insertText', { text: 'R' }, first.ticket)];
  session.gate.open();
  let statuses: string[] = [];
  let details: string[] = [];
  await act(async () => {
    statuses = await Promise.all(queued.map((entry) => entry.status));
    details = await Promise.all(queued.map((entry) => entry.detail));
  });
  console.log('G244-QUEUED statuses=' + JSON.stringify(statuses));
  console.log('G244-QUEUED details=' + JSON.stringify(details));
  expect(statuses.every((status) => !status.startsWith('rejected') && status !== 'refused' && status !== 'cancelled')).toBe(true);
  await session.checkpointTaken(1);
  const expected = session.text();
  const version = session.version();
  expect(expected).toContain('PQR');
  await session.crash();
  session.expectRecovered(expected);
  expect(session.version()).toBe(version);
}, TIMEOUT);

test('a long parsed-source session retains log bytes linearly', async () => {
  const parser = await createYrsSession({ clientId: 9299 });
  let document: Document;
  try {
    parser.openDocx(documentBytes(), true);
    document = parser.materializeDocx()!;
  } finally { parser.destroy(); }
  const worker = start();
  liveWorkers.push(worker);
  const client = new DocxSourceWorkerClient(worker);
  await client.openSource({ document });
  const initial = await client.commandInit({ session: 'parsed-log', incarnation: 1, epoch: 1, ...INIT });
  expect(initial.seed).toBeDefined();
  const retained = new Map<number, number>();
  const payloadBytes = (payload: unknown) => Object.values((payload ?? {}) as object).reduce((total: number, item: unknown) =>
    total + (item instanceof ArrayBuffer || ArrayBuffer.isView(item) ? item.byteLength : 0), 0);
  const session = createDocxWorkerEditSession<Document>({
    source: document, connection: await createDocxCommandConnection(client, initial),
    echo: { paint() {}, drop() {}, present() {} }, maxRecoveryAttempts: 1,
    commandBytes: (envelope) => {
      const bytes = payloadBytes(envelope.payload);
      retained.set(envelope.seq, bytes);
      return bytes;
    },
    recovery: { reopen: () => Promise.reject(new Error('Unexpected reopen')) },
  });
  const submit = async (kind: string, payload: Record<string, unknown>) =>
    ((await session.submit(kind as never, payload as never).ack) as { status: string }).status;
  const total = () => [...retained.values()].reduce((sum, bytes) => sum + bytes, 0);
  const samples: number[] = [];
  try {
    for (let round = 1; round <= 300; round += 1) {
      expect(await submit('insertText', { text: LETTERS[round % 26]! })).toBe('applied');
      expect(await submit('insertText', { text: 'x' })).toBe('applied');
      expect(await submit('deleteBackward', {})).toBe('applied');
      if (round % 100 === 0) samples.push(total());
    }
  } finally { await session.dispose().catch(() => {}); }
  console.log(`G244-PARSED-LOG bytes=${samples.join(',')}`);
  const [at100, at200, at300] = samples as [number, number, number];
  expect(at300 - at200).toBeLessThanOrEqual(1.5 * (at200 - at100));
}, 600_000);
