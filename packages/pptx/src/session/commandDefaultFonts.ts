import type { BundledFontFace, BundledFontLoader, BundledFontScript, ProgressiveFontFace, SliceLocation } from '@betteroffice/fonts';
import { importFonts, importProgressiveFonts } from '../../../../shared/office-session/fontModules';
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

/** @internal */
const DEFAULT_FONT_MAX_ATTEMPTS = 3;
/** @internal */
const DEFAULT_FONT_RETRY_DELAYS_MS = [2000, 8000];
const NASKH_FAMILIES = new Set(['simplified arabic', 'traditional arabic']);

let decoderReady: Promise<unknown> = Promise.resolve();

const sessionDecoder = sessionFontDecoder(async (bytes, options) => {
  await decoderReady.catch(() => {});
  return decodeFontFace(bytes, options);
});

function settledHttpMiss(error: unknown): boolean {
  const cause = isImplicitFontCdnError(error) ? error.cause : error;
  const status = (cause as { status?: number } | null)?.status;
  return status === 403 || status === 404 || status === 410;
}

const ignoreCdnFailure: FontCdnReporter = () => {};

function createFontSession() {
  const reported = new Set<FontCdnDiagnostic['pack']>();
  const report: FontCdnReporter = (error) => {
    if (reported.has(error.diagnostic.pack)) return;
    reported.add(error.diagnostic.pack);
    reportImplicitCdnFailure(error);
  };
  const slices = new Map<string, Promise<ProgressiveFontFace>>();
  const coverage = new WeakMap<BundledFontLoader, Promise<ArrayBuffer>>();
  return {
    report,
    loadCoverage(load: BundledFontLoader): Promise<ArrayBuffer> {
      let pending = coverage.get(load);
      if (!pending) {
        pending = load();
        coverage.set(load, pending);
        void pending.then(() => { coverage.delete(load); }, (error: unknown) => {
          if (!settledHttpMiss(error)) coverage.delete(load);
          if (isImplicitFontCdnError(error)) report(error);
        });
      }
      return pending;
    },
    openSlices(location: SliceLocation): Promise<ProgressiveFontFace> {
      const key = location.ref.index;
      let pending = slices.get(key);
      if (!pending) {
        pending = importProgressiveFonts().then(({ openSlicedFace }) =>
          openSlicedFace({ ...location, reportCdnFailure: ignoreCdnFailure }));
        slices.set(key, pending);
        void pending.then(() => { slices.delete(key); }, (error: unknown) => {
          if (!settledHttpMiss(error)) slices.delete(key);
          if (isImplicitFontCdnError(error)) report(error);
        });
      }
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
  async load(...configuration: [fontCdn?: FontCdnConfiguration]) {
    if (configuration.length) setFontCdnConfiguration(configuration[0]);
    const fonts = await importFonts();
    const session = createFontSession();
    const source = fonts.createFontProvider(undefined, {
      decoder: sessionDecoder, strict: 'except-optional-cdn-failure', reportCdnFailure: session.report,
    });
    let coverage: typeof source | undefined;
    const provider = { ...source, resolveScriptFallback(...args: Parameters<typeof source.resolveScriptFallback>) {
      coverage ??= fonts.createFontProvider(undefined, { decoder: sessionDecoder, strict: true });
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
  sliced = false,
): void {
  decoderReady = wasmReady;
  if (!scripts.length) return;
  void provider().then(async (source) => {
    for (const script of scripts) {
      const load = source.resolveScriptFallback(script as BundledFontScript, false, false);
      if (load?.slices && sliced && (await import('./commandSlicedFonts')).slicedFacesSupported()) {
        void fontSession(source).openSlices(load.slices).catch(() => {});
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
  const attempted = new Set<string>();
  const attempts = new Map<string, number>();
  const retryAt = new Map<string, number>();
  const pending = new Map<string, Promise<DefaultFontFace | undefined>>();
  const installations = new Map<string, Promise<boolean>>();
  const registered = new Map<string, boolean>();
  const preferredFaces = new WeakSet<DefaultFontFace>();
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

  const retryable = (id: string) => !attempted.has(id) ||
    ((attempts.get(id) ?? 0) < DEFAULT_FONT_MAX_ATTEMPTS && Date.now() >= (retryAt.get(id) ?? Infinity));
  const failed = (id: string) => {
    if (!retryAt.has(id)) retryAt.set(id,
      Date.now() + (DEFAULT_FONT_RETRY_DELAYS_MS[(attempts.get(id) ?? 1) - 1] ?? Infinity));
  };

  function publish(wanted: FontRequest[], faces: (DefaultFontFace | undefined)[]): DefaultFontLanding[] {
    if (disposed) return [];
    const additions: DefaultFontLanding[] = [];
    for (const [index, face] of faces.entries()) {
      if (!face || suppliedFamily(wanted[index])) continue;
      const registeredKey = faceKey(face, face.kind);
      const preferred = preferredFaces.has(face);
      if (registered.get(registeredKey) === true || (!preferred && registered.has(registeredKey))) continue;
      try {
        if (!prepared) {
          beforeRegister?.();
          prepared = true;
        }
        register(face);
        registered.set(registeredKey, preferred);
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
    const previousAttempts = attempts.get(id);
    const previousRetry = retryAt.get(id);
    attempted.add(id);
    attempts.set(id, (previousAttempts ?? 0) + 1);
    retryAt.delete(id);
    const retryCoverage = (error: { retryAt?: number }) => {
      attempts.set(id, previousAttempts ?? 0);
      retryAt.set(id, error.retryAt ?? Date.now() + DEFAULT_FONT_RETRY_DELAYS_MS[0]);
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
      let available = false;
      provider ??= Promise.resolve().then(() => shared ? shared() : commandFontProvider.load()).catch((error) => {
        provider = undefined;
        throw error;
      });
      void provider.then(async (source) => {
        if (disposed) return undefined;
        const session = fontSession(source);
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
                if (!settledHttpMiss(error) && !(await importFonts()).isMissingFontAddon(error)) {
                  if (isImplicitFontCdnError(error)) retryCoverage(error);
                  else failed(id);
                }
                return undefined;
              });
              if (bytes && await scriptFontFace('arabic', bold, italic, bytes, face.family)) {
                await waitForPrimary;
                preferred = { family: face.family, bold, italic, bytes };
                preferredFaces.add(preferred);
              }
            }
          }
          load = named ?? light ?? (preferred ? undefined : source.resolveLastResort(face.family, bold, italic, 'powerpoint'));
        }
        if (!load && !preferred) return undefined;
        available = true;
        let triedSlices = false;
        if (load?.slices && sliced?.supported()) {
          triedSlices = true;
          try {
            const opened = await session.openSlices(load.slices);
            if (disposed || suppliedFamily(face)) return undefined;
            const skeleton = face.kind === 'coverage' ? await scriptFontFace(face.script, bold, italic, opened.bytes(), face.requestedFamily)
              : { family: face.family, bold: face.bold, italic: face.italic, bytes: opened.bytes() };
            if (skeleton && !families.has(skeleton.family.trim().toLowerCase())) {
              const ensureText = (text: string) => opened.ensureText(text).catch((error: unknown) => {
                if (isImplicitFontCdnError(error)) session.report(error);
                throw error;
              });
              sliced.add(skeleton.family, { get revision() { return opened.revision; }, bytes: () => opened.bytes(),
                ensureText, prefetch: (text) => { void ensureText(text).catch(() => {}); },
                chunksOf: (text) => opened.chunksOf(text), loadedChunks: () => opened.loadedChunks() }, load, session.report);
              return disposed ? undefined : skeleton;
            }
          } catch (error) {
            if (isImplicitFontCdnError(error) && error.diagnostic.pack === 'latin') throw error;
          }
        }
        const loading = load && (face.kind === 'coverage' ? session.loadCoverage(load) : load());
        const bytes = preferred?.bytes ?? await loading?.catch(async (error: unknown) => {
          const implicit = isImplicitFontCdnError(error);
          if (implicit) session.report(error);
          if ((settledHttpMiss(error) && (face.kind === 'coverage' || (triedSlices && !implicit))) ||
              (face.kind === 'coverage' && !implicit && (await importFonts()).isMissingFontAddon(error))) {
            available = false;
            return undefined;
          }
          throw error;
        });
        if (!bytes || disposed || suppliedFamily(face)) return undefined;
        let loaded = preferred;
        if (lastResort) loaded = { family: lastResort.family, bold: lastResort.weight === 700,
          italic: lastResort.style === 'italic', bytes, kind: 'coverage' };
        else if (face.kind === 'coverage') loaded = await scriptFontFace(face.script, bold, italic, bytes, face.requestedFamily);
        else loaded ??= { family: face.family, bold: face.bold, italic: face.italic, bytes, ...(light ? { ownWeight: true } : {}) };
        if (!loaded || disposed || suppliedFamily(face)) return undefined;
        if (!preferred && registered.get(faceKey(loaded, loaded.kind)) === true) {
          available = false;
          return undefined;
        }
        sliced?.plain(loaded.family, loaded.bold, loaded.italic);
        const identity = JSON.stringify([loaded.family.trim().toLowerCase(), loaded.bold, loaded.italic, Boolean(preferred)]);
        let installation = installations.get(identity);
        if (!installation) {
          installation = Promise.resolve().then(async () => !disposed && (!install || (await install([loaded])).length > 0));
          installations.set(identity, installation);
          const current = installation;
          const forget = () => { if (installations.get(identity) === current) installations.delete(identity); };
          void installation.then((success) => { if (!success) forget(); }, forget);
        }
        return await installation && !disposed ? loaded : undefined;
      }).then((loaded) => {
        pending.delete(id);
        if (!loaded && available && !disposed && !suppliedFamily(face)) failed(id);
        if (finished && loaded) publish([face], [loaded]);
        else finish(loaded);
      }, (error: unknown) => {
        pending.delete(id);
        if (isImplicitFontCdnError(error)) {
          if (face.kind === 'coverage' || (face.kind === 'primary' && face.preferArabic)) {
            retryCoverage(error);
          } else if (previousAttempts === undefined) {
            attempted.delete(id);
            attempts.delete(id);
          } else {
            attempts.set(id, previousAttempts);
            if (previousRetry !== undefined) retryAt.set(id, previousRetry);
          }
        } else if (!disposed && !suppliedFamily(face)) failed(id);
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
        const naskh = NASKH_FAMILIES.has(family);
        const preferArabic = naskh && face.scripts?.includes('arabic') === true && !families.has(family);
        const primary: FontRequest = { kind: 'primary', family: face.family, bold: face.bold, italic: face.italic, preferArabic };
        wanted.set(key(primary), primary);
        for (const script of face.scripts ?? []) {
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
        (face.kind !== 'primary' || registered.get(faceKey(face)) !== true) &&
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
