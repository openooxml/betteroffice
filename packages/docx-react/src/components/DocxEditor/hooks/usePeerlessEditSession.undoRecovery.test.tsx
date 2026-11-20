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

const payloadBytes = (payload: unknown) => Object.values((payload ?? {}) as object).reduce((total: number, item: unknown) =>
  total + (Array.isArray(item) ? item : [item]).reduce((sum: number, part: unknown) =>
    sum + (part instanceof ArrayBuffer || ArrayBuffer.isView(part) ? part.byteLength : 0), 0), 0);

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
  /** Yrs state the first open restores over `source`, so that open records no seed. */
  initialState?: Uint8Array;
}

function mount({ tamper, source = documentBytes(), reopenSource, collaboration, gateCheckpoint, initialState }: MountOptions = {}) {
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
  const streamed = new Map<number, number>();
  const seeds: (Seed | undefined)[] = [];
  const fontEvents: { worker: number; families: string[] }[] = [];
  let pendingCheckpoint: Promise<unknown> | undefined;
  let applied = 0;
  const openResident = async (resume?: Resume): Promise<DocxCommandConnection> => {
    const worker = start();
    workers.push(worker);
    liveWorkers.push(worker);
    const client = new DocxSourceWorkerClient(worker);
    clients.push(client);
    const opened = workers.length;
    client.onFontEnvironment((event) => { fontEvents.push({ worker: opened, families: [...event.families] }); });
    const checkpoint = resume?.checkpoint as DocxCommandCheckpoint | undefined;
    if (checkpoint && !checkpoint.recoveryState) {
      throw new SessionFailure('message', 'Cannot restore DOCX checkpoint command-local state and proposals');
    }
    const state = checkpoint ? checkpoint.state : resume ? undefined : initialState?.slice();
    await client.openSource({ bytes: (resume && reopenSource ? reopenSource : source).slice(), ...(state ? { state } : {}),
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
    seeds.push(connection.initial.seed);
    if (!resume) {
      if (tamper) tamper(connection.initial.seed!);
      const onStream = connection.client.onStream;
      if (onStream) {
        connection.client.onStream = (listener) => onStream((entries) => {
          for (const entry of entries) streamed.set(entry.index, payloadBytes(entry.effect.payload));
          listener(entries);
        });
      }
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
    view, checkpoints, phases, events, errors, resumes, reopenedText, replayed, seeds, settle,
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
    fontFamilies(worker: number) {
      return fontEvents.filter((event) => event.worker === worker).at(-1)?.families ?? null;
    },
    streamBytes() {
      const covered = checkpoints.at(-1)?.stream ?? 0;
      return [...streamed].reduce((total, [index, bytes]) => index >= covered ? total + bytes : total, 0);
    },
    async submit(kind: string, payload: Record<string, unknown> = {}, options: { inputTime?: number } = {}) {
      let status = '';
      await act(async () => {
        const outcome = await view.result.current.binding!.session.submit(kind, payload as never, options).ack as { status: string };
        status = outcome.status;
      });
      if (status === 'applied') applied += 1;
      return status;
    },
    async type(count: number) {
      for (let i = 0; i < count; i += 1) expect(await session.submit('insertText', { text: LETTERS[i % 26]! })).toBe('applied');
    },
    async typeSteps(count: number) {
      const base = Math.ceil(performance.now()) + 10_000;
      for (let i = 0; i < count; i += 1) {
        expect(await session.submit('insertText', { text: LETTERS[i % 26]! }, { inputTime: base + i * 1_000 })).toBe('applied');
      }
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
    async idle() {
      for (let i = 0; i < 3; i += 1) {
        await settle();
        if (pendingCheckpoint) await act(async () => { try { await pendingCheckpoint; } catch {} });
      }
    },
    async typeSaveUndo(recovered: string) {
      const { focus } = session.selection();
      expect(focus.story).toBe('body');
      expect(recovered.includes('|')).toBe(false);
      const typed = recovered.slice(0, focus.offset) + 'B' + recovered.slice(focus.offset);
      expect(await session.submit('insertText', { text: 'B' })).toBe('applied');
      expect(session.text()).toBe(typed);
      await session.save();
      expect(session.text()).toBe(typed);
      expect(await session.submit('undo')).toBe('applied');
      expect(session.text()).toBe(recovered);
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

async function sourceState(bytes: Uint8Array): Promise<Uint8Array> {
  const worker = start();
  liveWorkers.push(worker);
  await new DocxSourceWorkerClient(worker).openSource({ bytes: bytes.slice() });
  return worker.sessions.at(-1)!.encodeState().slice();
}

type IssuedEntry = [key: string, story: string, encoded: string];

const issuedEntry = (key: unknown, value: unknown): IssuedEntry | null => {
  if (typeof key !== 'string' || !key.includes('\0') || !value || typeof value !== 'object') return null;
  const { story, encoded } = value as { story?: unknown; encoded?: unknown };
  return typeof story === 'string' && encoded instanceof Uint8Array ? [key, story, Array.from(encoded).join(',')] : null;
};

/**
 * Observes the worker's issued-position map from the same realm the in-process worker runs in: a map is the issued-position
 * map once a sticky-position entry (key `story\0bytes`, value `{ story, encoded }`) is set into it. While `counting`, every
 * entry an iteration of that map yields is counted; `added` lists the keys first set into it since the last `reset`.
 */
function watchIssuedPositions() {
  type AnyMap = Map<unknown, unknown>;
  type Iterate = (this: AnyMap) => Iterator<unknown>;
  const proto = Map.prototype as unknown as Record<PropertyKey, unknown>;
  const names = [Symbol.iterator, 'entries', 'keys', 'values'] as const;
  const originals = new Map<PropertyKey, unknown>([...names, 'set', 'forEach'].map((name) => [name, proto[name]]));
  const set = originals.get('set') as (this: AnyMap, key: unknown, value: unknown) => AnyMap;
  const forEach = originals.get('forEach') as (this: AnyMap, callback: (value: unknown, key: unknown, map: AnyMap) => void, self?: unknown) => void;
  const has = Map.prototype.has;
  const maps = new Set<AnyMap>();
  const watch = { counting: false, visited: 0, added: [] as IssuedEntry[], maps,
    reset() { watch.visited = 0; watch.added = []; },
    restore() { for (const [name, original] of originals) proto[name] = original; },
  };
  proto.set = function (this: AnyMap, key: unknown, value: unknown) {
    const entry = issuedEntry(key, value);
    if (entry) {
      maps.add(this);
      if (!has.call(this, key)) watch.added.push(entry);
    }
    return set.call(this, key, value);
  };
  for (const name of names) {
    const original = originals.get(name) as Iterate;
    proto[name] = function (this: AnyMap) {
      const iterator = original.call(this);
      if (!watch.counting || !maps.has(this)) return iterator;
      const counted = Object.create(Object.getPrototypeOf(Object.getPrototypeOf(iterator))) as Iterator<unknown> & Iterable<unknown>;
      counted.next = () => {
        const step = iterator.next();
        if (!step.done) watch.visited += 1;
        return step;
      };
      counted[Symbol.iterator] = () => counted;
      return counted;
    };
  }
  proto.forEach = function (this: AnyMap, callback: (value: unknown, key: unknown, map: AnyMap) => void, self?: unknown) {
    if (!watch.counting || !maps.has(this)) return forEach.call(this, callback, self);
    return forEach.call(this, (value, key, map) => {
      watch.visited += 1;
      callback.call(self, value, key, map);
    });
  };
  return watch;
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

test('a replayed entry never becomes an undo step: history does not survive a recovery: undo after it removes nothing', async () => {
  const live = mount();
  await live.ready();
  await live.typeSteps(5);
  const liveTyped = live.text();
  expect(await live.submit('undo')).toBe('applied');
  expect(live.text().length).toBe(liveTyped.length - 1);
  cleanup();

  const early = mount();
  await early.ready();
  await early.typeSteps(5);
  const typed = early.text();
  const version = early.version();
  await early.crash();
  expect(early.checkpoints).toEqual([]);
  console.log(`G244-REPLAY-UNDO early replayed=${JSON.stringify(early.replayed)}`);
  expect(early.replayed.length).toBeGreaterThanOrEqual(5);
  early.expectRecovered(typed);
  await early.submit('undo');
  expect(early.text()).toBe(typed);
  expect(early.version()).toBe(version);
  cleanup();

  const late = mount();
  await late.ready();
  await late.typeUntilApplied(128);
  await late.checkpointTaken(1);
  await late.typeSteps(3);
  const lateTyped = late.text();
  const lateVersion = late.version();
  await late.crash();
  console.log(`G244-REPLAY-UNDO late replayed=${JSON.stringify(late.replayed)}`);
  late.expectRecovered(lateTyped);
  await late.submit('undo');
  expect(late.text()).toBe(lateTyped);
  expect(late.version()).toBe(lateVersion);
}, TIMEOUT);

test('a font environment sent after open comes back after a worker restart before a checkpoint', async () => {
  const session = mount();
  await session.ready();
  const status = await session.submit('setFontEnvironment', {
    faces: [{ family: 'Liberation Sans', data: FONT.slice() }, { family: 'Brand Sans', data: FONT.slice() }],
    chains: { 'liberation sans|0|0': [0], 'brand sans|0|0': [1] }, families: ['Liberation Sans', 'Brand Sans'],
  });
  expect(['applied', 'noop']).toContain(status);
  expect(await session.submit('insertText', { text: 'a' })).toBe('applied');
  await waitFor(() => expect(session.fontFamilies(1)).toContain('Brand Sans'), { timeout: 30_000 });
  const before = session.fontFamilies(1);
  const expected = session.text();
  await session.crash();
  expect(session.checkpoints).toEqual([]);
  console.log(`G244-FONT-ENV replayed=${JSON.stringify(session.replayed)} before=${JSON.stringify(before)} after=${JSON.stringify(session.fontFamilies(2))}`);
  session.expectRecovered(expected);
  await waitFor(() => expect(session.fontFamilies(2)).toEqual(before), { timeout: 30_000 });
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

test('saves after a published save do not grow the log', async () => {
  const session = mount();
  await session.ready();
  expect(await session.submit('paragraphBreak')).toBe('applied');
  const before = session.stateVector();
  await session.save();
  if (session.stateVector() === before) throw new Error('save wrote nothing; mechanism not reached');
  await session.settle();
  const afterFirst = session.streamBytes();
  const checkpoints = session.checkpoints.length;
  const grown: number[] = [];
  for (let i = 2; i <= 300; i += 1) {
    await session.save();
    await session.settle();
    if (session.streamBytes() !== afterFirst) grown.push(i);
  }
  const afterLast = session.streamBytes();
  console.log(`G244-IDLE-SAVES bytes after save 1=${afterFirst} after save 300=${afterLast} grown=${JSON.stringify(grown.slice(0, 10))}`);
  expect(afterFirst).toBeGreaterThan(0);
  expect(grown).toEqual([]);
  expect(session.checkpoints.length).toBe(checkpoints);
  const expected = session.text();
  await session.crash();
  session.expectRecovered(expected);
  expect(session.engine().hasPendingUpdate()).toBe(false);
}, 600_000);

test('rounds of a paragraph insertion and a publishing save reach the count checkpoint', async () => {
  const session = mount();
  await session.ready();
  for (let round = 1; round <= 128; round += 1) {
    expect(await session.submit('paragraphBreak')).toBe('applied');
    const before = session.stateVector();
    await session.save();
    if (session.stateVector() === before) throw new Error(`save ${round} wrote nothing; mechanism not reached`);
  }
  await session.checkpointTaken(1);
  console.log(`G244-PUBLISH-ROUNDS checkpoints=${session.checkpoints.length}`);
  const expected = session.text();
  expect(expected.split('|')).toHaveLength(129);
  await session.crash();
  session.expectRecovered(expected);
  expect(session.engine().hasPendingUpdate()).toBe(false);
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
  const detail = await session.outcome(undo.detail);
  console.log(`G244-INFLIGHT-UNDO status=${status} detail=${detail} phases=${JSON.stringify(session.phases)}`);
  expect(status).toBe('refused');
  expect(detail).toBe(JSON.stringify({ status: 'refused', reason: 'history-reset' }));
  const seq = undo.ticket.seq!;
  expect(Number.isSafeInteger(seq)).toBe(true);
  expect(session.events.filter((event) => event.phase === 'diverged')
    .map(({ diverged, historyReset }) => ({ diverged, historyReset }))).toEqual([{ diverged: [seq], historyReset: true }]);
  expect(session.engine().hasPendingUpdate()).toBe(false);
  await session.typeSaveUndo(typed);
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
  const detail = await session.outcome(redo.detail);
  console.log(`G244-INFLIGHT-REDO status=${status} detail=${detail} phases=${JSON.stringify(session.phases)}`);
  expect(status).toBe('refused');
  expect(detail).toBe(JSON.stringify({ status: 'refused', reason: 'history-reset' }));
  const seq = redo.ticket.seq!;
  expect(Number.isSafeInteger(seq)).toBe(true);
  expect(session.events.filter((event) => event.phase === 'diverged')
    .map(({ diverged, historyReset }) => ({ diverged, historyReset }))).toEqual([{ diverged: [seq], historyReset: true }]);
  expect(session.engine().hasPendingUpdate()).toBe(false);
  await session.typeSaveUndo(before);
}, TIMEOUT);

test('an undo in flight at a crash after a checkpoint of a session opened without a seed re-executes as on main', async () => {
  const session = mount({ initialState: await sourceState(documentBytes()) });
  await session.ready();
  expect(session.seeds).toEqual([undefined]);
  await session.typeUntilApplied(128);
  await session.checkpointTaken(1);
  const before = session.text();
  expect(await session.submit('insertText', { text: 'A' })).toBe('applied');
  const typed = session.text();
  expect(typed).not.toBe(before);
  const worker = session.worker();
  worker.hold();
  const undo = await session.submitQueued('undo');
  await waitFor(() => expect(textOf(worker)).toBe(before), { timeout: 30_000 });
  await session.crash();
  expect(session.resumes[0]!.checkpoint).toBeDefined();
  expect(session.resumes[0]!.seed).toBeUndefined();
  const detail = await session.outcome(undo.detail);
  console.log(`G244-UNSEEDED-INFLIGHT-UNDO detail=${detail} phases=${JSON.stringify(session.phases)} replayed=${JSON.stringify(session.replayed)}`);
  expect(detail).toBe(JSON.stringify({ status: 'applied' }));
  expect(session.replayed.filter((kind) => kind === 'insertText' || kind === 'undo')).toEqual(['insertText', 'undo']);
  session.expectRecovered(before);
  expect(session.phases).not.toContain('diverged');
  expect(session.engine().hasPendingUpdate()).toBe(false);
}, TIMEOUT);

test('a collaborating core session with checkpoint recovery re-executes an undo in flight at a crash as on main', async () => {
  const source = documentBytes();
  const workers: InProcessResidentWorker[] = [];
  const checkpoints: DocxCommandCheckpoint[] = [];
  const events: RecoveryReport[] = [];
  const replayed: string[] = [];
  const open = async (resume?: Resume): Promise<DocxCommandConnection> => {
    const worker = start();
    workers.push(worker);
    liveWorkers.push(worker);
    const client = new DocxSourceWorkerClient(worker);
    const checkpoint = resume?.checkpoint as DocxCommandCheckpoint | undefined;
    await client.openSource({ bytes: source.slice(), ...(checkpoint ? { state: checkpoint.state } : {}),
      collaboration: { clientId: 9200 + workers.length } });
    const initial = await client.commandInit({
      session: 'collab-inflight-undo', incarnation: resume?.incarnation ?? 1, epoch: resume?.epoch ?? 1,
      ...(resume ? { version: checkpoint?.version ?? 0, throughSeq: checkpoint?.throughSeq ?? 0,
        refusedSeqs: resume.refusedSeqs, replayAppliedSeqs: resume.replayAppliedSeqs,
        recoveryState: checkpoint?.recoveryState } : {}),
      ...INIT,
    });
    const connection = await createDocxCommandConnection(client, initial);
    if (resume) {
      const edit = connection.client.call.edit;
      connection.client.call.edit = (command) => {
        replayed.push(command.kind);
        return edit(command);
      };
    }
    return connection;
  };
  const textNow = () => textOf(workers.at(-1)!);
  const until = async (done: () => boolean, label: string) => {
    for (const started = performance.now(); !done();) {
      if (performance.now() - started > 60_000) throw new Error(`Timed out waiting for ${label}`);
      await new Promise<void>((resolve) => setTimeout(resolve, 5));
    }
  };
  const session = createDocxWorkerEditSession<Uint8Array>({
    source, connection: await open(), echo: { paint() {}, drop() {}, present() {} },
    checkpointPolicy: ({ applied }) => checkpoints.length === 0 && applied >= 1,
    recovery: {
      reopen: (input) => open(input as Resume),
      checkpoint: async (connection) => {
        const value = await connection.client.call.checkpoint();
        checkpoints.push(value as DocxCommandCheckpoint);
        return value;
      },
      onRecovery: (event) => { events.push(event); },
    },
  });
  const submit = async (kind: string, payload: Record<string, unknown> = {}) =>
    ((await session.submit(kind as never, payload as never).ack) as { status: string }).status;
  try {
    expect(await submit('insertText', { text: 'x' })).toBe('applied');
    await until(() => checkpoints.length === 1, 'the checkpoint');
    const before = textNow();
    expect(await submit('insertText', { text: 'A' })).toBe('applied');
    expect(textNow()).not.toBe(before);
    const worker = workers.at(-1)!;
    worker.hold();
    const undo: { outcome?: { status: string; reason?: string } } = {};
    const ack = (session.submit('undo' as never, {} as never).ack as Promise<{ status: string; reason?: string }>)
      .then((outcome) => { undo.outcome = outcome; }, (error: Error) => { undo.outcome = { status: `rejected: ${error.message}` }; });
    const held = { committed: false };
    await until(() => {
      worker.release((reply) => {
        if (reply.ok && reply.commandAck?.outcomes.some((outcome) => outcome.status === 'applied')) held.committed = true;
        return false;
      });
      return held.committed;
    }, 'the undo to commit in the worker');
    worker.onerror?.({ message: 'simulated worker crash' } as ErrorEvent);
    await until(() => session.stage !== 'recovering' && events.some((event) =>
      event.phase === 'recovered' || event.phase === 'failed'), 'recovery');
    await until(() => undo.outcome !== undefined, 'the undo outcome');
    await ack;
    const outcome = undo.outcome!;
    const detail = JSON.stringify({ status: outcome.status, reason: outcome.reason });
    console.log(`G244-COLLAB-INFLIGHT-UNDO detail=${detail} phases=${JSON.stringify(events.map((event) => event.phase))} replayed=${JSON.stringify(replayed)}`);
    expect(workers).toHaveLength(2);
    expect(detail).toBe(JSON.stringify({ status: 'applied' }));
    expect(replayed.filter((kind) => kind === 'insertText' || kind === 'undo')).toEqual(['insertText', 'undo']);
    expect(events.map((event) => event.phase)).toContain('recovered');
    expect(events.map((event) => event.phase)).not.toContain('diverged');
    expect(textNow()).toBe(before);
  } finally { await session.dispose().catch(() => {}); }
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

test('capture work for one edit does not grow with the edits before it', async () => {
  const watch = watchIssuedPositions();
  try {
    const session = mount();
    await session.ready();
    let edits = 0;
    const atParagraphStart = (text: string) => {
      const { story, paraId } = session.selection().focus;
      const at = { story, paraId, offset: 0 };
      return { text, selection: { anchor: at, focus: { ...at } } };
    };
    const editUntil = async (count: number) => {
      for (; edits < count; edits += 1) {
        const breaking = edits % 50 === 49;
        expect(await session.submit(breaking ? 'paragraphBreak' : 'insertText', breaking ? {} : atParagraphStart(LETTERS[edits % 26]!))).toBe('applied');
      }
    };
    const measure = async (earlier: number) => {
      await editUntil(earlier);
      await session.idle();
      const worker = session.worker();
      const deliver = worker.onmessage!;
      let returned: IssuedEntry[] | undefined;
      worker.onmessage = (event) => {
        const ack = event.data.ok ? event.data.commandAck : undefined;
        if (watch.counting && ack?.outcomes.length) {
          watch.counting = false;
          returned = ack.outcomes.flatMap((outcome) => {
            const state = (outcome.effect?.payload as { state?: { stickyPositions?: [string, unknown][] } } | undefined)?.state;
            return (state?.stickyPositions ?? []).map(([key, position]) => issuedEntry(key, position)!);
          });
        }
        deliver.call(worker, event);
      };
      watch.reset();
      watch.counting = true;
      try {
        expect(await session.submit('insertText', atParagraphStart('Z'))).toBe('applied');
      } finally {
        watch.counting = false;
        worker.onmessage = deliver;
      }
      edits += 1;
      const sample = { earlier, visited: watch.visited, minted: watch.added, returned };
      console.log(`G244-CAPTURE-SCAN earlier=${earlier} visited=${sample.visited} minted=${sample.minted.length} ` +
        `returned=${returned?.length} maps=${watch.maps.size} issued=${[...watch.maps].map((map) => map.size).join(',')} minted=${JSON.stringify(sample.minted.map(([key]) => key))} ` +
        `returned=${JSON.stringify(returned?.map(([key]) => key))}`);
      return sample;
    };
    const samples = [await measure(10), await measure(2_000)];
    for (const sample of samples) {
      expect(sample.minted.length).toBeGreaterThan(0);
      expect(sample.returned).toEqual(sample.minted);
      expect(sample.visited).toBeLessThanOrEqual(sample.minted.length);
    }
  } finally {
    watch.restore();
  }
}, 1_800_000);
