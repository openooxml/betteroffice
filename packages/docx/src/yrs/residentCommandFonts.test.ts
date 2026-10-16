import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import type { ResidentCommandInitRequest } from './residentCommandProtocol';
import { collectCommandFonts } from './residentCommandFonts';

async function initializeFonts(
  init: ResidentCommandInitRequest,
  failure?: { base?: Error; substitute?: Error },
  progress = { layouts: 0, bindings: 0 },
) {
  const path = `${import.meta.dir}/residentEngineWorker.ts`;
  const source = readFileSync(path, 'utf8');
  const start = source.indexOf('  if (snapshot.fontsRevision !== fontsRevision) {');
  const end = source.indexOf('  for (const { story, env } of snapshot.renderInputs)', start);
  const initStart = source.indexOf("  if (request.type === 'commandInit') {");
  const initEnd = source.indexOf("  if (request.type === 'edit') {", initStart);
  const reservationStart = source.indexOf('const appliedLayoutReservations =');
  const reservations = source.slice(reservationStart, source.indexOf('\n/**', reservationStart));
  const reservationImport = source.match(/^import .* from '\.\/reservations';$/m)?.[0] ?? '';
  const fontImport = source.match(/^import .* from '\.\/residentCommandFonts';$/m)?.[0] ?? '';
  const hydrateSignature = source.slice(source.indexOf('function hydrate('), source.indexOf('): { layoutJson:', source.indexOf('function hydrate(')));
  const contents = `${fontImport}
    ${reservationImport}
    ${reservations}
    export async function initialize(session, request, bindResidentCommands) {
      let fontsRevision = -1;
      let glyphCache;
      const activeOffscreenPageIds = [], pendingOffscreenPageIds = new Set(), intactBackBuffers = new Map();
    const openedDocument = true, previewing = false, commandHost = null, commandHooksFactory = true, openedViaSource = false;
    let sourcePreviewing = false;
      const setFrameDisplayWindow = () => {}, replyCommand = () => {};
      let unsubscribe = null, pendingUpdates = [];
      ${hydrateSignature}): void {
        ${source.slice(start, end)}
        session.layoutDocumentWithRegionsRetained(snapshot.layoutInput);
      }
      ${source.slice(initStart, initEnd)}
    }
  `;
  const built = await Bun.build({
    entrypoints: [path], target: 'bun',
    plugins: [{ name: 'font-init-seam', setup(build) {
      build.onLoad({ filter: /\/residentEngineWorker\.ts$/ }, () => ({ contents, loader: 'ts' }));
    } }],
  });
  expect(built.success).toBe(true);
  const module = await import(`data:text/javascript;base64,${Buffer.from(await built.outputs[0].text()).toString('base64')}`);
  const substitutes: [number, string][] = [];
  let nextId = 40;
  let hydrated: unknown;
  let bound: ResidentCommandInitRequest | undefined;
  const session = {
    clearFonts() {},
    registerFont() {
      if (failure?.base) throw failure.base;
      return nextId++;
    },
    registerSubstituteFont(base: number, family: string) {
      if (failure?.substitute) throw failure.substitute;
      substitutes.push([base, family]);
      return family === 'Unknown' ? base : nextId++;
    },
    layoutDocumentWithRegionsRetained(input: string) { progress.layouts += 1; hydrated = JSON.parse(input); },
  };
  await module.initialize(session, { type: 'commandInit', id: 1, request: init }, (value: ResidentCommandInitRequest) => {
    progress.bindings += 1;
    bound = value;
    return { initial: async () => ({}) };
  });
  return { substitutes, hydrated, bound };
}

test('registration_slots_remap_substitute_aliases', async () => {
  const fontChains = { base: [0], alias: [1, 2, 3], nested: [4, 1], duplicate: [0, 1, 2] };
  const remapped = { base: [40], alias: [40, 41, 42], nested: [43, 40], duplicate: [40, 41] };
  for (const useSlotChains of [false, true]) {
    const init: ResidentCommandInitRequest & { fontRegistrationSlots: true | { fontChains: Record<string, number[]> } } = {
      session: 'fonts', incarnation: 1, epoch: 1, readOnly: false,
      fonts: [Uint8Array.of(1), { substituteOf: 0, family: 'Unknown' }, Uint8Array.of(2),
        { substituteOf: 2, family: 'Calibri' }, { substituteOf: 1, family: 'Arial' }],
      fontRegistrationSlots: useSlotChains ? { fontChains } : true,
      layoutInput: JSON.stringify({ measurement: { fontChains: useSlotChains ? { base: [0], alias: [0, 1] } : fontChains } }),
      layoutWithRegions: true, layoutRevision: 1, extras: JSON.stringify({ fontChains }),
    };
    const { substitutes, hydrated, bound } = await initializeFonts(init);
    expect(substitutes).toEqual([[40, 'Unknown'], [41, 'Calibri'], [40, 'Arial']]);
    expect(hydrated).toEqual({ measurement: { fontChains: remapped } });
    expect(JSON.parse(bound!.layoutInput)).toEqual(hydrated);
    expect(JSON.parse(bound!.extras)).toEqual({ fontChains: remapped });
    expect(init.fonts[1]).toEqual({ substituteOf: 0, family: 'Unknown' });
  }
});

test('collected_buffers_dispose_on_abort', async () => {
  const controller = new AbortController();
  const collected = collectCommandFonts(controller.signal);
  const bytes = Uint8Array.of(11, 12);
  const base = collected.engine.registerFont(bytes);
  const alias = collected.engine.registerSubstituteFont!(base, 'Calibri');
  expect([base, alias]).toEqual([0, 1]);
  expect(collected.fonts).toEqual([bytes, { substituteOf: base, family: 'Calibri' }]);
  expect(collected.fonts[0]).not.toBe(bytes);
  controller.abort();
  expect(collected.fonts).toEqual([]);
  expect(() => collected.engine.registerFont(bytes)).toThrow();
  expect(() => collected.engine.registerSubstituteFont!(base, 'Arial')).toThrow();
  collected.dispose();
  expect(collected.fonts).toEqual([]);
});

test('collected_buffers_dispose_without_abort', () => {
  const collected = collectCommandFonts();
  const base = collected.engine.registerFont(Uint8Array.of(1));
  collected.engine.registerSubstituteFont!(base, 'Calibri');
  expect(collected.fonts).toHaveLength(2);
  collected.dispose();
  expect(collected.fonts).toEqual([]);
  expect(() => collected.engine.registerFont(Uint8Array.of(2))).toThrow();
  expect(() => collected.engine.registerSubstituteFont!(base, 'Arial')).toThrow();
  collected.engine.clearFonts();
  expect(collected.fonts).toEqual([]);
});

test('command_init_rethrows_font_registration_traps_before_layout', async () => {
  for (const registration of ['base', 'substitute'] as const) {
    const trap = new WebAssembly.RuntimeError('unreachable');
    const progress = { layouts: 0, bindings: 0 };
    const init: ResidentCommandInitRequest = {
      session: 'fonts', incarnation: 1, epoch: 1, readOnly: false,
      fonts: [Uint8Array.of(1), { substituteOf: 0, family: 'Calibri' }],
      fontRegistrationSlots: true,
      layoutInput: JSON.stringify({ measurement: { fontChains: { face: [0, 1] } } }),
      layoutWithRegions: true, layoutRevision: 1, extras: '{}',
    };
    await expect(initializeFonts(init, { [registration]: trap }, progress)).rejects.toBe(trap);
    expect(progress).toEqual({ layouts: 0, bindings: 0 });
  }
});
