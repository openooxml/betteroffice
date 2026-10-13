import { expect, mock, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { createDocxWorkerEditSession } from '@betteroffice/docx/peerless';
import type { ResidentEngineWorkerClient, YrsResidentFontRegistration } from '@betteroffice/docx/yrs';
import { applyFrameDelta, decodeFrameDelta } from '@betteroffice/docx/layout/render';
import type { DocxCommandAck, DocxCommandEnvelope, ResidentCommandConnection } from '@betteroffice/docx/peerless';
import type { RustTextEngine } from '@betteroffice/docx/layout/measure/rustMeasureSource';
import { buildCommandFrame } from '../../../../../docx/src/yrs/residentCommandFrame';
import { registerCommandFonts } from '../../../../../docx/src/yrs/residentCommandFonts';
import { stageCommandPresentation } from '../../../../../docx/src/yrs/residentCommandPresentation';
import { openCommandEngine, closeCommandEngines } from '../../../../../docx/src/yrs/residentCommandTestSupport';
import { projectCommandEchoContext } from '../../../../../docx/src/yrs/residentCommandViewProjection';
import { fontEnvironmentForRegistrations } from './workerFonts';
import { createEchoPainter, type PeerlessEchoIntent } from './echoPainter';
import { createFramePresentation } from './framePresentation';
import { admitInputIntents, createInputAdmission } from './inputAdmission';
import { normalizeInputEvent } from './inputEvents';
import { createPeerlessViewStore } from './viewStore';

type ResidentCommandInitRequest = Parameters<ResidentEngineWorkerClient['commandInit']>[0];

async function openFontConnection(hold = false, fontKeys: readonly string[] = ['face']) {
  const source = readFileSync(`${import.meta.dir}/PeerlessDocxEditor.tsx`, 'utf8');
  const start = source.indexOf('    signal.throwIfAborted();', source.indexOf('const openConnection ='));
  const end = source.indexOf('\n  }, [controller]);', start);
  const transpiler = new Bun.Transpiler({ loader: 'tsx' });
  const body = transpiler.transformSync(`async function open(input, signal) {${source.slice(start, end).replace(
    'const fonts: YrsResidentFontRegistration[] = [];',
    'const fonts: YrsResidentFontRegistration[] = []; observeFonts(fonts);',
  )} }`);
  const registerFont = mock(() => 70);
  const registerSubstituteFont = mock((base: number) => base);
  const getRustTextEngine = mock(async () => ({ registerFont, registerSubstituteFont, clearFonts() {} }));
  const controller = new AbortController();
  let destroyed = false;
  let cleared = false;
  let collector: RustTextEngine | undefined;
  let collected: YrsResidentFontRegistration[] | undefined;
  let init: ResidentCommandInitRequest | undefined;
  let ready: () => void = () => {};
  const initialized = new Promise<void>((resolve) => { ready = resolve; });
  const initialize = async (value: ResidentCommandInitRequest) => {
    init = value;
    ready();
    if (hold) await new Promise((_, reject) => controller.signal.addEventListener('abort', () => reject(new Error('aborted')), { once: true }));
  };
  const client = {
    onFontEnvironment: () => () => {}, onMemory: () => () => {},
    prewarm: async () => {}, openSource: async () => ({ document: { package: { document: { comments: [] }, settings: {} } } }), fontRequirements: async () => '[]',
    async commandInit(value: ResidentCommandInitRequest) {
      await initialize(value);
      return { session: 'fonts', incarnation: 1, epoch: 1 };
    },
    destroy() { destroyed = true; },
  };
  const connection = { initial: { session: 'fonts', incarnation: 1, epoch: 1 }, client: { dispose: async () => client.destroy() } };
  const createResidentCommandConnection = mock(async (_client: unknown, value: ResidentCommandInitRequest | typeof connection.initial) => {
    if ('fonts' in value) await initialize(value);
    return connection;
  });
  const dependencies = {
    active: { current: null },
    takePreloadedResidentEngineWorker: () => null,
    controller: { fontChoices() {} }, fontSource: { current: null }, fontAbort: { current: null }, fontJob: { current: null },
    fontEnvironmentForRegistrations,
    worker: { current: null },
    sourceComments: { current: [] },
    renderEnvironment: { current: { theme: null, requested: '' } },
    ResidentEngineWorkerClient: function () { return client; },
    decodeDocxHostJson: () => ({ document: { package: { document: { comments: [] }, settings: {} } } }),
    buildResidentRegionLayoutRequest: () => ({}),
    getRustTextEngine,
    observeFonts: (fonts: YrsResidentFontRegistration[]) => { collected = fonts; },
    createRustMeasureSource: ({ engine }: { engine: RustTextEngine }) => {
      collector = engine;
      let ids: number[] = [];
      return {
        setCompat() {}, setEmbeddedFaces() {},
        async prepareFontRequirements() {
          ids = [engine.registerFont(Uint8Array.of(11, 12))];
          ids.push(engine.registerSubstituteFont!(ids[0], 'Calibri'));
          ids.push(engine.registerFont(Uint8Array.of(21, 22)));
        },
        measurementConfigForRequirements: () => ({ fontChains: Object.fromEntries(fontKeys.map((key) => [key, ids])) }),
        clear() { cleared = true; },
      };
    },
    extractEmbeddedFontFaces: async () => [],
    createResidentCommandConnection,
    discardCommandPresentation: () => {},
    settings: { current: { readOnly: false } }, identity: {}, retained: {}, pageIndexes: {}, prepared: {},
  };
  const open = new Function(...Object.keys(dependencies), `${body}\nreturn open;`)(...Object.values(dependencies)) as
    (bytes: Uint8Array, signal: AbortSignal) => Promise<typeof connection>;
  const result = open(Uint8Array.of(1), controller.signal);
  await initialized;
  return { result, controller, get init() { return init; }, registerFont, registerSubstituteFont,
    getRustTextEngine, get collector() { return collector; }, get collected() { return collected; },
    cleared: () => cleared, destroyed: () => destroyed };
}

test('repeated_open_allocates_no_main_fonts', async () => {
  for (let cycle = 0; cycle < 2; cycle += 1) {
    const opened = await openFontConnection();
    const connection = await opened.result;
    expect(opened.cleared()).toBe(true);
    expect(opened.collected).toEqual([]);
    expect(() => opened.collector!.registerFont(Uint8Array.of(1))).toThrow();
    await connection.client.dispose();
    opened.controller.abort();
    expect(opened.registerFont).not.toHaveBeenCalled();
    expect(opened.registerSubstituteFont).not.toHaveBeenCalled();
    expect(opened.getRustTextEngine).not.toHaveBeenCalled();
    expect(opened.destroyed()).toBe(true);
  }
});

test('command_init_receives_collected_font_slots', async () => {
  const opened = await openFontConnection();
  const connection = await opened.result;
  const init: Parameters<ResidentEngineWorkerClient['commandInit']>[0] = opened.init!;
  expect(init.fonts).toEqual([Uint8Array.of(11, 12), { substituteOf: 0, family: 'Calibri' }, Uint8Array.of(21, 22)]);
  expect(init.fontRegistrationSlots).toEqual({ fontChains: { face: [0, 1, 2] } });
  await connection.client.dispose();
  opened.controller.abort();
});

test('command_init_receives_display_font_slots', async () => {
  const opened = await openFontConnection();
  const connection = await opened.result;
  try {
    expect(JSON.parse(opened.init!.extras).fontChains).toEqual({ face: [0, 1, 2] });
  } finally {
    await connection.client.dispose();
    opened.controller.abort();
  }
});

test('typed_input_draws_on_page_canvas_with_authoritative_frame_held', async () => {
  const opened = await openFontConnection(false, ['calibri|0|0', 'liberation sans|0|0']);
  const openedConnection = await opened.result;
  const cleanups: (() => void | Promise<void>)[] = [];
  try {
    const init = opened.init!;
    const arm = await openCommandEngine(['abcd']);
    cleanups.push(closeCommandEngines);
    const font = new Uint8Array(readFileSync(new URL('../../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)));
    arm.engine.clearFonts();
    const registered = registerCommandFonts(arm.engine, init.fonts.map((item) => item instanceof Uint8Array ? font : item));
    const slots = init.fontRegistrationSlots;
    if (!slots || typeof slots !== 'object') throw new Error('Font registration slots missing');
    arm.engine.layoutDocumentWithRegionsRetained(registered.remap(JSON.stringify({
      bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] }, renderEnv: {},
      measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' }, ...JSON.parse(init.layoutInput).measurement },
    }), slots.fontChains));
    const bytes = arm.engine.buildDisplayListFrame(registered.remap(init.extras, slots.fontChains), 0);
    const delta = decodeFrameDelta(bytes);
    const retained = applyFrameDelta(null, delta);
    const layout = { frame: retained, layoutRevision: init.layoutRevision };
    const echoContext = projectCommandEchoContext({ engine: arm.engine, selection: arm.state.selection, layout });
    const bitmap = (width: number, height: number, text: string) => ({ width, height, text, close() {} }) as unknown as ImageBitmap;
    const staged = await stageCommandPresentation({ layout, pixelRatio: 1, echoContext,
      pages: retained.pages.map(({ pageId, pageIndex }) => ({ pageId: pageId.toString(), pageIndex })),
      async rasterizePage() { return bitmap(retained.pages[0]!.page.width, retained.pages[0]!.page.height, 'abcd'); },
      async rasterizeBackground({ rect }) { return bitmap(rect.width, rect.height, ''); },
    });
    const candidate = buildCommandFrame({ frame: bytes, metadata: delta, coversThroughSeq: 0,
      caret: arm.engine.residentCaretSnapshot(), layoutRevision: init.layoutRevision,
      echoContext, stagedPresentation: staged.value,
    }).value;
    const identity = { session: init.session, incarnation: init.incarnation, epoch: init.epoch };
    const initial: ResidentCommandConnection['initial'] = { ...identity, version: 0, throughSeq: 0, outcomes: [],
      selection: arm.state.selection,
      editState: { readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null,
        canUndo: false, canRedo: false, hyperlink: null, tableContext: null },
      frame: candidate,
    };
    let pixels = '';
    const texts: string[] = [];
    const canvas = { width: retained.pages[0]!.page.width, height: retained.pages[0]!.page.height,
      getContext: () => ({ save() {}, restore() {}, setTransform() {}, transform() {}, beginPath() {}, rect() {}, clip() {}, fillRect() {},
        clearRect() { pixels = ''; }, drawImage(image: ImageBitmap & { text: string }) { pixels = image.text; },
        fillText(text: string) { texts.push(text); pixels = text; }, measureText: (text: string) => ({ width: text.length * 10 }),
        getImageData: () => ({ width: canvas.width, height: canvas.height }) as ImageData, putImageData() {},
      }),
    } as unknown as HTMLCanvasElement;
    const viewStore = createPeerlessViewStore({ identity, version: 0, throughSeq: 0, presentedThroughSeq: -1,
      workerSelection: initial.selection, anchors: null, editState: initial.editState, geometry: null });
    cleanups.push(() => viewStore.dispose());
    const canvasForPage = (page: { pageId: string }) => page.pageId === retained.pages[0]!.pageId.toString() ? canvas : null;
    const painter = createEchoPainter({ canvasForPage });
    cleanups.push(() => painter.dispose());
    const presentation = createFramePresentation({ viewStore, echoPainter: painter, canvasForPage,
      requestFullFrame() { throw new Error('Unexpected full frame request'); },
      prepareGeometry(frame) { return { identity, coversThroughSeq: frame.coversThroughSeq, frameEpoch: frame.revision,
        layoutRevision: frame.payload.layoutRevision, caret: null, selectionRects: [],
        pages: retained.pages.map(({ pageId, pageIndex, page }) => ({ page: { pageId: pageId.toString(), pageIndex },
          bounds: { x: 0, y: 0, width: page.width, height: page.height }, transform: [1, 0, 0, 1, 0, 0] as const })),
      }; },
    });
    cleanups.push(() => presentation.dispose());
    const commands: DocxCommandEnvelope[] = [];
    let acknowledge: ((ack: DocxCommandAck) => void) | undefined;
    const queued: (() => void)[] = [];
    const connection: ResidentCommandConnection = { initial, requestFullFrame: async () => initial,
      client: { failure: undefined, onFailure: () => () => {}, dispose: async () => {},
        call: { checkpoint: async () => ({ state: new Uint8Array(), epoch: 1, version: 0, throughSeq: 0 }),
          edit(command) { commands.push(command); return new Promise((resolve) => { acknowledge = resolve; }); },
        },
      },
    };
    const session = createDocxWorkerEditSession<Uint8Array, PeerlessEchoIntent>({ source: Uint8Array.of(1), connection,
      echo: presentation.hooks, maxRecoveryAttempts: 0, checkpointPolicy: () => false, now: () => 0,
      enqueue(callback) { queued.push(callback); },
      recovery: { reopen: async () => { throw new Error('Unexpected recovery'); } },
    });
    cleanups.push(() => session.dispose());
    const admission = createInputAdmission({ session, viewStore, echoPainter: painter });
    cleanups.push(() => admission.dispose());
    expect(pixels).toBe('abcd');
    const normalized = normalizeInputEvent({ source: 'beforeinput', inputType: 'insertText', data: 'A',
      isComposing: false, eventId: 1, inputTime: 0 }, { composing: false, compositionId: null,
      lastHandledEventId: null, trailingCompositionText: null });
    admitInputIntents(admission, normalized.intents, { source: 'typing', fence: normalized.fence });
    const ticket = admission.getSnapshot().tail!;
    while (queued.length) queued.shift()!();
    expect(commands.map(({ kind, payload }) => ({ kind, payload }))).toEqual([{ kind: 'insertText', payload: { text: 'A' } }]);
    expect(texts).toEqual(['Aabcd']);
    expect(pixels).toBe('Aabcd');
    acknowledge!({ ...initial, frame: undefined, version: 1, throughSeq: ticket.seq!,
      outcomes: [{ seq: ticket.seq!, status: 'applied' }] });
    expect((await ticket.ack).status).toBe('applied');
    await Promise.resolve();
    expect(session.state.throughSeq).toBe(ticket.seq!);
    expect(viewStore.getSnapshot().geometry?.coversThroughSeq).toBe(0);
    expect(session.state.presentedThroughSeq).toBe(0);
    expect(pixels).toBe('Aabcd');
    expect(texts).toEqual(['Aabcd']);
  } finally {
    for (const cleanup of cleanups.reverse()) await cleanup();
    await openedConnection.client.dispose();
    opened.controller.abort();
  }
}, 20_000);

test('editor_collected_buffers_dispose_on_abort', async () => {
  const opened = await openFontConnection(true);
  expect(opened.init!.fonts.length).toBeGreaterThan(0);
  opened.controller.abort();
  expect(opened.collected).toEqual([]);
  await expect(opened.result).rejects.toThrow();
  expect(opened.cleared()).toBe(true);
  expect(() => opened.collector!.registerFont(Uint8Array.of(1))).toThrow();
  expect(opened.destroyed()).toBe(true);
});
