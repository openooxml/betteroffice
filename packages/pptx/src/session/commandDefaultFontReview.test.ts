import { afterEach, expect, spyOn, test } from 'bun:test';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { isClientMessage, isHostMessage } from '../../../../shared/office-session/protocol';
import { defaultFontInstaller, type DefaultFontInstallation, type DefaultFontPayload } from '../../../../shared/pptxDefaultFonts';
import * as progressive from '@betteroffice/fonts/progressive';
import * as wasm from '../wasm/loader';
import * as requests from '../wasm/fontRequests';
import { createPresentationCommandClient, createSharedPresentationCommandClient } from './commandClient';
import { commandDefaultFontModule } from './commandHost';
import * as fontModule from './commandDefaultFonts';
import { createPresentationSessionHost } from './host';
import type { PresentationCommandClient, PresentationCommandIntent } from './commandProtocol';

const restorers: (() => void)[] = [];
const clients: PresentationCommandClient[] = [];
const regular = { family: 'Calibri', bold: false, italic: false };
const bytes = Uint8Array.of(1, 2, 3).buffer;

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

function deadline() {
  const timers: (() => void)[] = [];
  const original = globalThis.setTimeout;
  const timer = spyOn(globalThis, 'setTimeout').mockImplementation(((callback: () => void, delay?: number) => {
    if (delay === 5000) timers.push(callback);
    return original(callback, delay === 5000 ? 60_000 : delay);
  }) as typeof setTimeout);
  restorers.push(() => timer.mockRestore());
  return () => { for (const callback of timers.splice(0)) callback(); };
}

function connection(initial = ['active', 'target']) {
  let ids = [...initial];
  let version = 0;
  let next = ['target'];
  const registered: string[] = [];
  const faces: { bytes: Uint8Array | ArrayBuffer }[] = [];
  const replaced: string[] = [];
  const pngFonts: string[] = [];
  const changed = deferred<void>();
  const receipt = () => {
    ids = [...next]; version++; changed.resolve();
    return { applied: true, changedSlideIds: ids, deckSizeChanged: false, slideOrderChanged: true };
  };
  const handle = {
    version: () => `v${version}`, onUpdate: () => () => {}, setUndoCaptureMode() {},
    readCommandContent: () => ({ ok: true, version: `v${version}`, slides: ids.map((id) => ({ id })), stories: [], sources: [] }),
    encodeStateAsUpdate: () => Uint8Array.of(0), isProposalsAvailable: () => false,
    canUndo: () => true, canRedo: () => true, addUndoBoundary() {},
    historyStep: receipt, applyUpdate: receipt, deleteSlide: receipt,
    snapshot: () => ({ slides: ids.map((id) => ({ id, shapes: [] })) }),
    registerFont: (face: { family: string; bytes: Uint8Array | ArrayBuffer }) => { registered.push(face.family); faces.push(face); return registered.length; },
    replaceExportFace: (family: string) => { replaced.push(family); },
    renderPng: async (_: unknown, options: { fonts?: { family: string }[] }) => {
      pngFonts.push(...(options.fonts ?? []).map((face) => face.family));
      return Uint8Array.of(7);
    },
    dispose() {},
  } as unknown as wasm.PresentationHandle;
  const open = spyOn(wasm, 'openPresentation').mockReturnValue(handle);
  const metadata = spyOn(wasm, 'presentationMetadata').mockImplementation(() => ({
    size: { width: 100, height: 80 }, slides: ids.map((id, index) => ({ id, index, name: id, layoutPartPath: null })),
  }));
  const layout = spyOn(wasm, 'presentationDisplayListJson').mockImplementation(() => JSON.stringify({
    width: 100, height: 80, primitives: [], families: [...registered],
  }));
  const coverage = spyOn(requests, 'presentationEnableBrowserCoverageFallback').mockImplementation(() => {});
  const collect = spyOn(requests, 'presentationRequestedFontFaces').mockImplementation((_, id) =>
    id === 'active' ? [] : [{ ...regular, family: 'Cambria' }]);
  const provider = spyOn(fontModule.commandFontProvider, 'load').mockResolvedValue({
    resolve: () => async () => bytes, resolveFamily: () => undefined, resolveScriptFallback: () => undefined,
    resolveLastResort: () => { throw new Error('Unexpected substitution'); },
  });
  const fetch = spyOn(globalThis, 'fetch').mockRejectedValue(new Error('Network disabled'));
  restorers.push(() => { open.mockRestore(); metadata.mockRestore(); layout.mockRestore(); coverage.mockRestore(); collect.mockRestore(); provider.mockRestore(); fetch.mockRestore(); });
  const pair = createInProcessPair();
  const messages: unknown[] = [];
  createPresentationSessionHost({ ...pair.host, post(message, transfer) {
    messages.push(message); pair.host.post(message, transfer);
  } }, { initWasm: async () => {} });
  restorers.push(() => pair.client.close());
  return { transport: pair.client, messages, registered, faces, replaced, pngFonts, changed: changed.promise, collect, layout, coverage,
    replace(value: string[]) { next = value; } };
}

afterEach(async () => {
  await Promise.all(clients.splice(0).map((client) => client.dispose()));
  for (const restore of restorers.splice(0).reverse()) restore();
});

for (const failure of ['reject', 'timeout'] as const) test(`browser ${failure} publishes fallback layout without registering the face`, async () => {
  const expire = deadline();
  const wire = connection(['target']);
  const installing = deferred<void>();
  const completed = deferred<void>();
  const installation = Object.assign({ session: 'fonts', incarnation: 'worker', fonts: [] }, {
    [defaultFontInstaller]: async () => {
      installing.resolve();
      if (failure === 'reject') throw new Error('Browser refused font');
      await completed.promise;
    },
  });
  const opening = createPresentationCommandClient(Uint8Array.of(0), installation, wire.transport);
  await installing.promise;
  expect(wire.registered).toEqual([]);
  expect(wire.coverage).not.toHaveBeenCalled();
  expect(wire.layout).not.toHaveBeenCalled();
  if (failure === 'timeout') expire();
  const client = await opening;
  clients.push(client);
  expect(wire.registered).toEqual([]);
  expect(JSON.parse(new TextDecoder().decode(client.state.frame!.displayList)).families).toEqual([]);
  completed.resolve();
  await client.flush();
  expect(wire.registered).toEqual([]);
  expect(wire.coverage).not.toHaveBeenCalled();
});

test('only successful browser faces enter worker layout through the private handshake', async () => {
  const wire = connection(['target']);
  wire.collect.mockReturnValue([regular, { ...regular, family: 'Cambria' }]);
  const installed = deferred<void>();
  const installing = deferred<void>();
  const installation = Object.assign({ session: 'fonts', incarnation: 'worker', fonts: [] }, {
    [defaultFontInstaller]: async (faces: readonly (typeof regular & { bytes: ArrayBuffer })[]) => {
      installing.resolve(); await installed.promise;
      return faces.filter((face) => face.family === 'Calibri');
    },
  });
  const opening = createPresentationCommandClient(Uint8Array.of(0), installation, wire.transport);
  await installing.promise;
  expect(wire.registered).toEqual([]);
  expect(wire.layout).not.toHaveBeenCalled();
  installed.resolve();
  const client = await opening;
  clients.push(client);
  expect(wire.registered).toEqual(['Calibri']);
  expect(wire.coverage).toHaveBeenCalledTimes(1);
  expect(JSON.parse(new TextDecoder().decode(client.state.frame!.displayList)).families).toEqual(['Calibri']);
  expect(bytes.byteLength).toBe(3);
});

test('app registration overtakes a pending browser default without blocking its reply', async () => {
  const wire = connection();
  const installed = deferred<void>();
  const installing = deferred<void>();
  let allowed: ((family: string) => boolean) | undefined;
  const options = Object.assign({ session: 'fonts', incarnation: 'worker', fonts: [] }, {
    [defaultFontInstaller]: async (_faces, current) => { allowed = current; installing.resolve(); await installed.promise; },
  } satisfies DefaultFontInstallation);
  const client = await createPresentationCommandClient(Uint8Array.of(0), options, wire.transport);
  clients.push(client);
  await installing.promise;
  expect(allowed?.('Cambria')).toBe(true);
  const ack = await client.dispatch({ kind: 'registerFont', payload: { inputTime: 1, undoGroup: 'app-font',
    font: { family: 'Cambria', bytes: Uint8Array.of(9).buffer } } });
  expect(ack.outcomes[0].status).toBe('applied');
  expect(allowed?.(' CAMBRIA ')).toBe(false);
  installed.resolve();
  await client.dispatch({ kind: 'activateSlide', payload: { slideId: 'target' } });
  expect(wire.registered).toEqual(['Cambria']);
  expect(wire.coverage).not.toHaveBeenCalled();
});

test('expired module imports acknowledge open and cannot create a late manager', async () => {
  const expire = deadline();
  const wire = connection(['target', 'other']);
  const imported = deferred<typeof fontModule>();
  const importing = deferred<void>();
  const load = spyOn(commandDefaultFontModule, 'load').mockImplementation(() => { importing.resolve(); return imported.promise; });
  const create = spyOn(fontModule, 'createCommandDefaultFonts');
  restorers.push(() => { load.mockRestore(); create.mockRestore(); });
  const opening = createPresentationCommandClient(Uint8Array.of(0), { session: 'fonts', incarnation: 'worker', fonts: [] }, wire.transport);
  await importing.promise;
  expire();
  const client = await opening;
  clients.push(client);
  expect(client.state.frame).not.toBeNull();
  imported.resolve(fontModule);
  await Promise.resolve(); await Promise.resolve();
  await client.dispatch({ kind: 'activateSlide', payload: { slideId: 'other' } });
  expect(create).not.toHaveBeenCalled();
  expect(wire.registered).toEqual([]);
  expect(load).toHaveBeenCalledTimes(1);
});

for (const shared of [false, true]) for (const kind of ['undo', 'redo', 'applyRemoteUpdate', 'invoke'] as const) {
  test(`${kind} gates an implicit active slide change (${shared ? 'shared' : 'direct'})`, async () => {
    const wire = connection();
    const installed = deferred<void>();
    const installing = deferred<void>();
    const create = shared ? createSharedPresentationCommandClient : createPresentationCommandClient;
    const options = Object.assign({ session: 'fonts', incarnation: 'worker', fonts: [] }, {
      [defaultFontInstaller]: async () => { installing.resolve(); await installed.promise; },
    } satisfies DefaultFontInstallation);
    const client = await create(Uint8Array.of(0), options, wire.transport);
    clients.push(client);
    await installing.promise;
    if (kind === 'undo' || kind === 'redo') wire.replace(['active', 'restored']);
    const replies: string[] = [];
    const off = wire.transport.listen((message) => {
      if (isHostMessage(message) && message.kind === 'reply') replies.push('reply');
    });
    const input = { inputTime: 1, undoGroup: 'history' };
    const intent: PresentationCommandIntent = kind === 'invoke'
      ? { kind, payload: { ...input, method: 'deleteSlide', args: ['active'] } }
      : kind === 'applyRemoteUpdate' ? { kind, payload: { ...input, bytes: Uint8Array.of(1).buffer } }
      : { kind, payload: input };
    let published = false;
    const navigation = client.dispatch(intent).then((ack) => { published = true; return ack; });
    await wire.changed;
    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(published).toBe(false);
    expect(replies).toEqual([]);
    expect(client.state.activeSlideId).toBe('active');
    installed.resolve();
    const ack = await navigation;
    expect(ack.frame?.slideId).toBe(kind === 'undo' || kind === 'redo' ? 'restored' : 'target');
    expect((ack as DefaultFontPayload).defaultFonts?.map((face) => face.family)).toEqual(['Cambria']);
    off();
  });
}

test('opting out of default fonts opens with no default font protocol traffic', async () => {
  const wire = connection();
  const importing = spyOn(commandDefaultFontModule, 'load');
  restorers.push(() => importing.mockRestore());
  const sent: unknown[] = [];
  const client = await createPresentationCommandClient(Uint8Array.of(0), { session: 'fonts', incarnation: 'worker',
    fonts: [{ ...regular, bytes }], defaultFonts: false },
  { ...wire.transport, post(message, transfer) { sent.push(message); wire.transport.post(message, transfer); } });
  clients.push(client);
  await client.dispatch({ kind: 'activateSlide', payload: { slideId: 'target' } });
  await client.flush();
  expect(importing).not.toHaveBeenCalled();
  expect(wire.coverage).not.toHaveBeenCalled();
  expect(wire.collect).not.toHaveBeenCalled();
  expect(wire.registered).toEqual([]);
  expect(sent.every((message) => isClientMessage(message))).toBe(true);
  expect(wire.messages.every((message) => !isHostMessage(message) || message.kind !== 'event')).toBe(true);
});

test('opting out with no app fonts loads nothing', async () => {
  const wire = connection(['target']);
  const importing = spyOn(commandDefaultFontModule, 'load');
  restorers.push(() => importing.mockRestore());
  const client = await createPresentationCommandClient(Uint8Array.of(0),
    { session: 'fonts', incarnation: 'worker', fonts: [], defaultFonts: false }, wire.transport);
  clients.push(client);
  await client.flush();
  expect(importing).not.toHaveBeenCalled();
  expect(wire.collect).not.toHaveBeenCalled();
  expect(wire.registered).toEqual([]);
});

test('app supplied families keep their faces while other document families load defaults', async () => {
  const wire = connection(['target']);
  wire.collect.mockImplementation(() => [
    regular, { ...regular, bold: true }, { ...regular, family: 'Times New Roman' }, { ...regular, family: 'MS Mincho', italic: true },
  ]);
  const resolved: string[] = [];
  const provider = spyOn(fontModule.commandFontProvider, 'load').mockResolvedValue({
    resolve: (family: string) => { resolved.push(family); return async () => bytes; },
    resolveFamily: () => undefined, resolveScriptFallback: () => undefined,
    resolveLastResort: () => { throw new Error('Unexpected substitution'); },
  });
  restorers.push(() => provider.mockRestore());
  const faces: string[] = [];
  const options = Object.assign({ session: 'fonts', incarnation: 'worker', fonts: [{ family: ' calibri ', bold: false, italic: false, bytes }] }, {
    [defaultFontInstaller]: async (additions: readonly { family: string }[]) => { faces.push(...additions.map((face) => face.family)); },
  } satisfies DefaultFontInstallation);
  const client = await createPresentationCommandClient(Uint8Array.of(0), options, wire.transport);
  clients.push(client);
  await client.flush();
  expect(resolved.sort()).toEqual(['MS Mincho', 'Times New Roman']);
  expect(faces.sort()).toEqual(['MS Mincho', 'Times New Roman']);
  expect(wire.registered.sort()).toEqual(['MS Mincho', 'Times New Roman']);
  expect(wire.coverage).toHaveBeenCalledTimes(1);
});

test('invalid default font options are refused', async () => {
  const wire = connection();
  await expect(createPresentationCommandClient(Uint8Array.of(0),
    { session: 'fonts', incarnation: 'worker', fonts: [], defaultFonts: 'no' as unknown as boolean }, wire.transport)).rejects.toThrow();
});

function slicedCapableBrowser() {
  const original = Object.getOwnPropertyDescriptor(globalThis, 'navigator');
  Object.defineProperty(globalThis, 'navigator', { configurable: true, value: {
    userAgent: 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36' } });
  const open = spyOn(progressive, 'openSlicedFace').mockRejectedValue(new Error('Network disabled'));
  const load = Object.assign(async () => bytes, { slices: { ref: { format: 1, index: 'x', indexSha256: '' } } });
  spyOn(fontModule.commandFontProvider, 'load').mockResolvedValue({
    resolve: () => load, resolveFamily: () => undefined, resolveScriptFallback: () => undefined,
    resolveLastResort: () => { throw new Error('Unexpected substitution'); },
  } as never);
  restorers.push(() => {
    open.mockRestore();
    if (original) Object.defineProperty(globalThis, 'navigator', original);
    else Reflect.deleteProperty(globalThis, 'navigator');
  });
  return open;
}

async function openWith(wire: ReturnType<typeof connection>, options: { slicedFonts?: boolean }) {
  const sent: unknown[] = [];
  const client = await createPresentationCommandClient(Uint8Array.of(0), Object.assign({ session: 'fonts', incarnation: 'worker',
    fonts: [], ...options }, { [defaultFontInstaller]: async () => {} }),
  { ...wire.transport, post(message, transfer) { sent.push(message); wire.transport.post(message, transfer); } });
  clients.push(client);
  await client.flush();
  return { client, sent };
}

test('a session that does not ask for sliced faces loads, registers and exports whole faces, with nothing new on the wire', async () => {
  const wire = connection(['target']);
  const open = slicedCapableBrowser();
  const { client, sent } = await openWith(wire, {});
  expect(new Uint8Array(await client.renderPng!({ slide: 0 }))).toEqual(Uint8Array.of(7));
  expect(open).not.toHaveBeenCalled();
  expect(wire.registered).toEqual(['Cambria']);
  expect(wire.faces.map((face) => [...new Uint8Array(face.bytes)])).toEqual([[1, 2, 3]]);
  expect(wire.replaced).toEqual([]);
  expect(wire.pngFonts).toEqual([]);
  expect(JSON.stringify([sent, wire.messages]).toLowerCase()).not.toContain('sliced');
  expect(JSON.stringify([sent, wire.messages])).not.toContain('CoverageFont');
});

test('a session that asks for sliced faces tries them first', async () => {
  const wire = connection(['target']);
  const open = slicedCapableBrowser();
  await openWith(wire, { slicedFonts: true });
  expect(open).toHaveBeenCalledTimes(1);
  expect(wire.registered).toEqual(['Cambria']);
});

test('PNG export of a sliced family carries the whole face by id and by family, for text measured with browser fonts', async () => {
  const wire = connection(['target']);
  slicedCapableBrowser().mockResolvedValue({ revision: 1, bytes: () => bytes, ensureText: async () => 1, prefetch() {} } as never);
  const { client } = await openWith(wire, { slicedFonts: true });
  await client.renderPng!({ slide: 0 });
  expect(wire.replaced).toEqual(['Cambria']);
  expect(wire.pngFonts).toEqual(['Cambria']);
});

test('a sliced face opened on a later slide is upgraded for PNG export', async () => {
  const wire = connection(['active', 'target']);
  slicedCapableBrowser().mockResolvedValue({ revision: 1, bytes: () => bytes, ensureText: async () => 1, prefetch() {} } as never);
  const { client } = await openWith(wire, { slicedFonts: true });
  await client.dispatch({ kind: 'activateSlide', payload: { slideId: 'target' } });
  await client.flush();
  await client.renderPng!({ slide: 0 });
  expect(wire.replaced.filter((family) => family === 'Cambria').length).toBeGreaterThan(0);
});

