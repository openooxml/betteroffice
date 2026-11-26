import type { BundledFontFace, BundledFontLoader, BundledFontScript, ProgressiveFontFace, SliceLocation } from '@betteroffice/fonts';
import { importFonts, importProgressiveFonts } from '../../../../shared/office-session/fontModules';
import { slicedFacesSupported } from '../../../../shared/office-session/render/pptxSlicedFonts';
import { setFontCdnConfiguration, type FontCdnConfiguration } from '../../../../shared/office-session/fontCdn';
import { isImplicitFontCdnError, reportImplicitCdnFailure, type FontCdnDiagnostic, type FontCdnReporter } from '../../../../shared/fontCdnDiagnostics';
import { scriptFontFace } from './commandScriptFonts';
import type { DefaultFontFace, DefaultFontLanding, RequestedFontFace } from '../../../../shared/pptxDefaultFonts';
import { decodeFontFace } from '../wasm/loader';
import { sessionFontDecoder } from './sessionFontDecoder';
import type { createSlicedFontFaces } from './commandSlicedFonts';

type FontRequest =
  | { kind: 'primary'; family: string; bold: boolean; italic: boolean; preferArabic?: boolean }
  | { kind: 'last-resort'; family: string; bold: boolean; italic: boolean }
  | { kind: 'coverage'; script: string; requestedFamily?: string; bold: boolean; italic: boolean };

type PendingSliceOpen = {
  pending: Promise<ProgressiveFontFace | undefined>;
  callers?: (() => boolean)[];
};

/** @internal */
const DEFAULT_FONT_MAX_ATTEMPTS = 3;
/** @internal */
const DEFAULT_FONT_RETRY_DELAYS_MS = [2000, 8000];
const NASKH_FAMILIES = new Set(['simplified arabic', 'traditional arabic']);
export const UNCOVERED_EDITOR_SCRIPTS = new Set(['arabic']);

let decoderReady: Promise<unknown> = Promise.resolve();

const sessionDecoder = sessionFontDecoder(async (bytes, options) => {
  await decoderReady.catch(() => {});
  return decodeFontFace(bytes, options);
});

function httpStatus(error: unknown): number | undefined {
  const cause = isImplicitFontCdnError(error) ? error.cause : error;
  return (cause as { status?: number } | null)?.status;
}

async function settledFontMiss(error: unknown): Promise<boolean> {
  const status = httpStatus(error);
  if (status === 403 || status === 404 || status === 410) return true;
  return !isImplicitFontCdnError(error) && await importFonts().then(
    ({ isMissingFontAddon }) => isMissingFontAddon(error), () => false);
}

function createFontSession() {
  const reported = new Set<FontCdnDiagnostic['pack']>();
  const blockedPacks = new Map<FontCdnDiagnostic['pack'], Error>();
  const report: FontCdnReporter = (error) => {
    const { pack } = error.diagnostic;
    if (isImplicitFontCdnError(error) && httpStatus(error) === 403) blockedPacks.set(pack, error);
    if (reported.has(pack)) return;
    reported.add(pack);
    reportImplicitCdnFailure(error);
  };
  const ignoreCdnFailure: FontCdnReporter = () => {};
  const slices = new Map<string, PendingSliceOpen>();
  const coverage = new WeakMap<BundledFontLoader, Promise<ArrayBuffer>>();
  return {
    report,
    checkImplicitCdnPack(pack: FontCdnDiagnostic['pack']): void {
      const error = blockedPacks.get(pack);
      if (error) throw error;
    },
    substitutes: new WeakMap<ArrayBuffer, unknown>(),
    loadCoverage(load: BundledFontLoader): Promise<ArrayBuffer> {
      let pending = coverage.get(load);
      if (!pending) {
        pending = load();
        coverage.set(load, pending);
        void pending.then(() => { coverage.delete(load); }, async (error: unknown) => {
          if (isImplicitFontCdnError(error)) report(error);
          if (!(await settledFontMiss(error))) coverage.delete(load);
        });
      }
      return pending;
    },
    openSlices(location: SliceLocation, current = () => true): Promise<ProgressiveFontFace | undefined> {
      const key = location.ref.index;
      const existing = slices.get(key);
      if (existing) {
        existing.callers?.push(current);
        return existing.pending;
      }
      const callers = [current];
      const scopedLocation = { ...location, reportCdnFailure: ignoreCdnFailure };
      const pending = importProgressiveFonts().then(({ openSlicedFace }) => {
        opening.callers = undefined;
        if (callers.some((isCurrent) => isCurrent())) return openSlicedFace(scopedLocation);
        slices.delete(key);
      });
      const opening: PendingSliceOpen = { pending, callers };
      slices.set(key, opening);
      void pending.then((opened) => { if (opened) slices.delete(key); }, async (error: unknown) => {
        opening.callers = undefined;
        if (isImplicitFontCdnError(error)) report(error);
        if (!(await settledFontMiss(error))) slices.delete(key);
      });
      return pending;
    },
  };
}

const fontSessions = new WeakMap<object, ReturnType<typeof createFontSession>>();
function fontSession(source: object) {
  let session = fontSessions.get(source);
  if (!session) {
    session = createFontSession();
    fontSessions.set(source, session);
  }
  return session;
}

export const commandFontProvider = {
  /** @internal */
  async load(
    ...configuration: [fontCdn?: FontCdnConfiguration]
  ): Promise<ReturnType<typeof import('@betteroffice/fonts').createFontProvider>> {
    if (configuration.length) setFontCdnConfiguration(configuration[0]);
    const fonts = await importFonts();
    const session = createFontSession();
    const runtime = { decoder: sessionDecoder, checkImplicitCdnPack: session.checkImplicitCdnPack };
    const primaryRuntime = {
      ...runtime, strict: 'except-optional-cdn-failure' as const, reportCdnFailure: session.report,
      onSubstitute: (bytes: ArrayBuffer, error: unknown) => { session.substitutes.set(bytes, error); },
    };
    const source = fonts.createFontProvider(undefined, primaryRuntime);
    let coverage: typeof source | undefined;
    const provider = { ...source, resolveScriptFallback(...args: Parameters<typeof source.resolveScriptFallback>) {
      coverage ??= fonts.createFontProvider(undefined, { ...runtime, strict: true });
      return coverage.resolveScriptFallback(...args);
    } };
    fontSessions.set(provider, session);
    return provider;
  },
};

/**
 * Starts loading the regular coverage face of each script before the deck is parsed. `provider` returns the
 * session's one font provider, the one later given to `createCommandDefaultFonts`, so a face started here is
 * fetched and decoded once. Decoding waits for `wasmReady`, where the font decoder lives.
 */
export function prefetchCoverageFaces(
  scripts: readonly string[], wasmReady: Promise<unknown>, provider: () => ReturnType<typeof commandFontProvider.load>,
  sliced = false, current = () => true,
): void {
  decoderReady = wasmReady;
  const coverage = scripts.filter((script) => !UNCOVERED_EDITOR_SCRIPTS.has(script));
  if (!coverage.length) return;
  void provider().then(async (source) => {
    for (const script of coverage) {
      if (!current()) return;
      const load = source.resolveScriptFallback(script as BundledFontScript, false, false);
      if (load?.slices && sliced && slicedFacesSupported()) {
        void fontSession(source).openSlices(load.slices, current).catch(() => {});
      } else if (load) void fontSession(source).loadCoverage(load).catch(() => {});
    }
  }).catch(() => {});
}

export function createCommandDefaultFonts(
  supplied: readonly { family: string }[],
  collect: (slideId?: string) => RequestedFontFace[],
  register: (face: DefaultFontFace) => void,
  settled: (faces: DefaultFontLanding[]) => void,
  install?: (faces: readonly DefaultFontFace[]) => Promise<readonly DefaultFontFace[]>,
  beforeRegister?: () => void,
  shared?: () => ReturnType<typeof commandFontProvider.load>,
  sliced?: ReturnType<typeof createSlicedFontFaces>,
) {
  const families = new Set(supplied.map((face) => face.family.trim().toLowerCase()));
  const attempts = new Map<string, { count: number; retryAt?: number }>();
  const pending = new Map<string, Promise<DefaultFontFace | undefined>>();
  const installations = new Map<string, Promise<boolean>>();
  const registered = new Map<string, 'substitute' | 'regular' | 'preferred'>();
  const faceStatus = new WeakMap<DefaultFontFace, 'substitute' | 'preferred'>();
  let provider: ReturnType<typeof commandFontProvider.load> | undefined;
  let disposed = false;
  let prepared = false;
  const cancellations = new Set<() => void>();
  const key = (face: FontRequest) => JSON.stringify([face.kind,
    face.kind === 'coverage' ? [face.script, face.requestedFamily?.trim().toLowerCase()] : face.family.trim().toLowerCase(),
    face.bold, face.italic, face.kind === 'primary' && face.preferArabic === true]);
  const faceKey = ({ family, bold, italic }: Omit<DefaultFontFace, 'bytes'>, kind = 'primary') =>
    JSON.stringify([kind, family.trim().toLowerCase(), bold, italic]);
  const suppliedFamily = (face: FontRequest) => face.kind === 'primary' && families.has(face.family.trim().toLowerCase());

  const retryable = (id: string) => {
    const attempt = attempts.get(id);
    return !attempt || (attempt.count < DEFAULT_FONT_MAX_ATTEMPTS && Date.now() >= (attempt.retryAt ?? Infinity));
  };
  const failed = (id: string) => {
    const attempt = attempts.get(id);
    if (attempt) attempt.retryAt ??= Date.now() + (DEFAULT_FONT_RETRY_DELAYS_MS[attempt.count - 1] ?? Infinity);
  };

  function publish(wanted: FontRequest[], faces: (DefaultFontFace | undefined)[]): DefaultFontLanding[] {
    if (disposed) return [];
    const additions: DefaultFontLanding[] = [];
    for (const [index, face] of faces.entries()) {
      if (!face || suppliedFamily(wanted[index])) continue;
      const registeredKey = faceKey(face, face.kind);
      const status = faceStatus.get(face) ?? 'regular';
      const previous = registered.get(registeredKey);
      if (previous === status || previous === 'preferred' || (previous && status === 'substitute')) continue;
      try {
        if (!prepared) {
          beforeRegister?.();
          prepared = true;
        }
        register(face);
        registered.set(registeredKey, status);
        additions.push({ family: face.family, bold: face.bold, italic: face.italic,
          ...(face.kind ? { kind: face.kind } : {}), ...(face.ownWeight ? { ownWeight: true } : {}) });
      } catch { failed(key(wanted[index])); }
    }
    if (additions.length) settled(additions);
    return additions;
  }

  function request(face: FontRequest): Promise<DefaultFontFace | undefined> {
    const id = key(face);
    const existing = pending.get(id);
    if (existing) return existing;
    if (!retryable(id) || suppliedFamily(face)) return Promise.resolve(undefined);
    const previousAttempts = attempts.get(id)?.count ?? 0;
    const attempt: { count: number; retryAt?: number } = { count: previousAttempts + 1 };
    attempts.set(id, attempt);
    let reportCdnFailure: FontCdnReporter | undefined;
    const recordFailure = async (error: unknown): Promise<'miss' | 'retry'> => {
      const cdn = isImplicitFontCdnError(error) ? error : undefined;
      if (cdn) reportCdnFailure?.(cdn);
      if (await settledFontMiss(error)) return 'miss';
      if (!disposed && !suppliedFamily(face)) {
        if (cdn) {
          attempt.count = previousAttempts;
          const retryAt = cdn.retryAt ?? Date.now() + DEFAULT_FONT_RETRY_DELAYS_MS[0];
          attempt.retryAt = Math.max(attempt.retryAt ?? 0, retryAt);
        } else failed(id);
      }
      return 'retry';
    };
    const task = new Promise<DefaultFontFace | undefined>((resolve) => {
      let finished = false;
      const finish = (value?: DefaultFontFace) => {
        if (finished) return;
        finished = true;
        clearTimeout(timer);
        cancellations.delete(cancel);
        resolve(value);
      };
      const cancel = () => finish();
      const timer = setTimeout(cancel, 5000);
      cancellations.add(cancel);
      provider ??= Promise.resolve().then(() => shared ? shared() : commandFontProvider.load()).catch((error) => {
        provider = undefined;
        throw error;
      });
      void provider.then(async (source) => {
        if (disposed) return undefined;
        const session = fontSession(source);
        reportCdnFailure = session.report;
        const waitForPrimary = face.kind === 'primary' && face.preferArabic
          ? pending.get(key({ ...face, preferArabic: false })) : undefined;
        let bold = face.bold;
        let italic = face.italic;
        let load: BundledFontLoader | undefined;
        let light: (() => Promise<ArrayBuffer>) | undefined;
        let preferred: DefaultFontFace | undefined;
        let lastResort: BundledFontFace | undefined;
        if (face.kind === 'coverage') {
          const script = face.script as BundledFontScript;
          const family: [string?] = face.requestedFamily === undefined ? [] : [face.requestedFamily];
          load = source.resolveScriptFallback(script, bold, italic, ...family);
          if (!load && (bold || italic)) {
            bold = false;
            italic = false;
            load = source.resolveScriptFallback(script, bold, italic, ...family);
          }
        } else if (face.kind === 'last-resort') {
          const { resolveLastResortChoice } = await importFonts();
          lastResort = resolveLastResortChoice(face.family, bold, italic, 'powerpoint').face;
          load = source.resolveLastResort(face.family, bold, italic, 'powerpoint');
        } else {
          const named = source.resolve(face.family, bold, italic) ?? source.resolveFamily(face.family, bold, italic);
          light = named ? undefined : source.resolveLightVariant?.(face.family, bold, italic);
          if (!named && !light && face.preferArabic) {
            const arabic = source.resolveScriptFallback('arabic', bold, italic, face.family);
            if (arabic) {
              const bytes = await session.loadCoverage(arabic).catch(async (error: unknown) => {
                await recordFailure(error);
                return undefined;
              });
              const arabicFace = bytes && await scriptFontFace('arabic', bold, italic, bytes, face.family);
              if (bytes && arabicFace) {
                await waitForPrimary;
                preferred = { family: face.family, bold, italic, bytes,
                  ...(bold !== arabicFace.bold || italic !== arabicFace.italic ? { ownWeight: true } : {}) };
                faceStatus.set(preferred, 'preferred');
              }
            }
          }
          load = named ?? light ?? (preferred ? undefined : source.resolveLastResort(face.family, bold, italic, 'powerpoint'));
        }
        if (!load && !preferred) return undefined;
        if (load?.slices && sliced?.supported()) {
          try {
            const opened = await session.openSlices(load.slices);
            if (!opened || disposed || suppliedFamily(face)) return undefined;
            const skeleton = face.kind === 'coverage' ? await scriptFontFace(face.script, bold, italic, opened.bytes(), face.requestedFamily)
              : { family: face.family, bold: face.bold, italic: face.italic, bytes: opened.bytes() };
            if (skeleton && !families.has(skeleton.family.trim().toLowerCase())) {
              const ensureText = (text: string) => opened.ensureText(text).catch((error: unknown) => {
                if (isImplicitFontCdnError(error)) session.report(error);
                throw error;
              });
              const whole = Object.assign(load.bind(undefined), { slices: load.slices, reportCdnFailure: session.report });
              sliced.add(skeleton.family, { get revision() { return opened.revision; }, bytes: () => opened.bytes(),
                ensureText, prefetch: (text) => { void ensureText(text).catch(() => {}); },
                chunksOf: (text) => opened.chunksOf(text), loadedChunks: () => opened.loadedChunks() }, whole);
              return disposed ? undefined : skeleton;
            }
          } catch (error) {
            if (isImplicitFontCdnError(error) && error.diagnostic.pack === 'latin') throw error;
          }
        }
        const loading = load && (face.kind === 'coverage' ? session.loadCoverage(load) : load());
        const bytes = preferred?.bytes ?? await loading;
        if (!bytes || disposed || suppliedFamily(face)) return undefined;
        let loaded = preferred;
        if (lastResort) loaded = { family: lastResort.family, bold: lastResort.weight === 700,
          italic: lastResort.style === 'italic', bytes, kind: 'coverage' };
        else if (face.kind === 'coverage') loaded = await scriptFontFace(face.script, bold, italic, bytes, face.requestedFamily);
        else loaded ??= { family: face.family, bold: face.bold, italic: face.italic, bytes, ...(light ? { ownWeight: true } : {}) };
        if (!loaded) {
          if (!disposed && !suppliedFamily(face)) failed(id);
          return undefined;
        }
        const substitute = session.substitutes.get(bytes);
        const substituteFailure = substitute === undefined ? undefined : await recordFailure(substitute);
        if (disposed || suppliedFamily(face)) return undefined;
        const previous = registered.get(faceKey(loaded, loaded.kind));
        if (!preferred && previous === 'preferred') return undefined;
        if (!load?.slices || substituteFailure !== 'retry') sliced?.plain(loaded.family, loaded.bold, loaded.italic);
        if (substituteFailure) {
          if (previous) return undefined;
          faceStatus.set(loaded, 'substitute');
        }
        const identity = JSON.stringify([loaded.family.trim().toLowerCase(), loaded.bold, loaded.italic,
          Boolean(preferred), substitute !== undefined]);
        let installation = installations.get(identity);
        if (!installation) {
          installation = Promise.resolve().then(async () => !disposed && (!install || (await install([loaded])).length > 0));
          installations.set(identity, installation);
          const current = installation;
          const forget = () => { if (installations.get(identity) === current) installations.delete(identity); };
          void installation.then((success) => { if (!success || substitute !== undefined) forget(); }, forget);
        }
        const installed = await installation;
        if (disposed || suppliedFamily(face)) return undefined;
        if (!installed) failed(id);
        return installed ? loaded : undefined;
      }).then((loaded) => {
        pending.delete(id);
        if (finished && loaded) publish([face], [loaded]);
        else finish(loaded);
      }, async (error: unknown) => {
        await recordFailure(error);
        pending.delete(id);
        finish();
      });
    });
    pending.set(id, task);
    return task;
  }

  function requests(slideId?: string): FontRequest[] {
    if (disposed) return [];
    try {
      const wanted = new Map<string, FontRequest>();
      for (const face of collect(slideId)) {
        const family = face.family.trim().toLowerCase();
        const scripts = face.scripts?.filter((script) => !UNCOVERED_EDITOR_SCRIPTS.has(script));
        const naskh = NASKH_FAMILIES.has(family);
        const preferArabic = naskh && scripts?.includes('arabic') === true && !families.has(family);
        const primary: FontRequest = { kind: 'primary', family: face.family, bold: face.bold, italic: face.italic, preferArabic };
        wanted.set(key(primary), primary);
        for (const script of scripts ?? []) {
          if (script === 'arabic' && preferArabic) {
            const fallback: FontRequest = { kind: 'last-resort', family: face.family, bold: face.bold, italic: face.italic };
            wanted.set(key(fallback), fallback);
            continue;
          }
          const coverage: FontRequest = { kind: 'coverage', script, bold: face.bold, italic: face.italic,
            requestedFamily: script === 'arabic' ? face.family : undefined };
          wanted.set(key(coverage), coverage);
        }
      }
      return [...wanted.values()].filter((face) => !suppliedFamily(face) &&
        (face.kind !== 'primary' || registered.get(faceKey(face)) !== 'preferred') &&
        (retryable(key(face)) || pending.has(key(face))));
    } catch { return []; }
  }

  return {
    isLoading(): boolean { return pending.size > 0; },
    hasPending(slideId?: string): boolean { return requests(slideId).length > 0; },
    load(slideId?: string): DefaultFontLanding[] | Promise<DefaultFontLanding[]> {
      const wanted = requests(slideId);
      if (!wanted.length) return [];
      return Promise.all(wanted.map(request)).then((faces) => publish(wanted, faces));
    },
    supply(family: string): void { families.add(family.trim().toLowerCase()); },
    dispose(): void {
      disposed = true;
      for (const cancel of cancellations) cancel();
      pending.clear();
      installations.clear();
    },
  };
}
